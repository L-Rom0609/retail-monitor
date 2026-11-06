# -*- coding: utf-8 -*-
"""
即时零售数据监控看板 —— 一键数据更新脚本（新手友好）
==================================================
最简单用法（推荐）：
    把【订单底表.csv】放进桌面文件夹「数据看板--底表更新」，双击「一键更新看板.bat」。

★ 归属表唯一真源（2026-11-06 用户规则）★
    仓名-负责人归属关系表一律以仓库内这份为准，且只能改这一份：
        配置与映射/仓名-负责人归属关系表.xlsx
    脚本不再读取桌面或其它位置的归属表副本（历史上两版并存曾导致 30 个仓的历史被删）。

其它用法（可选）：
    1) 双击本脚本运行                                  → 自动扫描底表文件夹
    2) python update_dashboard.py 订单.csv              → 手动指定订单底表
    3) 把 csv 拖到 .bat 图标上放开                      → 用拖进来的订单底表（xlsx 参数已废弃）

★ 历史累积存档（重要）★
    本脚本带「历史数据仓库」：每次更新都是往数据仓库里【追加/覆盖】，
    而不是把看板整体替换掉。
      · 放入 10 月的表 → 更新 → 放回 9 月的表（10 月的删了）→ 更新
        → 看板里 9 月和 10 月的数据都还在，不会丢。
      · 同一个月的表反复更新 → 按原样覆盖，不会重复累加成两倍。
      · 文件夹里可以同时放多份 csv（比如 9 月、10 月各一份），一次全部读进来。
    数据仓库文件：历史数据存档.json（请勿删除，它就是看板的历史数据本体）

    如需清空重来（丢弃全部历史）：python update_dashboard.py --replace
    只按某个月重算（例）：      python update_dashboard.py --drop 2026-09

口径（2026-10-29 与需求方确认）：
  1. 平台仅保留：美团闪购 / 淘宝闪购零售 / 京东秒送（其他平台完全排除）
  2. 订单状态全部计入（含已取消 / 待配送），不过滤
  3. 营收 = W 列「用户实付金额」之和
  4. 配送平台(AC列) = 自配送 → 补单；其余（含空值）→ 真实订单
     （特例：SELF_DELIVERY_REAL_STORES 内仓的自配送 = 真实订单，见下方配置）
  5. 配送门店(AA列) 匹配不上归属表的订单直接剔除，不计入任何指标
  6. 仓名 / 负责人取归属表 B / C 列（按 AA=A 列关联）
  7. ★ 永久过滤：渠道店铺ID（A^ 列）命中 EXCLUDE_STORE_IDS 的订单整单剔除，
     不计入任何指标（规则见下方「EXCLUDE_STORE_IDS」）

重要：脚本只读取底表，绝不修改或删除您的原始文件。

环境：Python 3.8+，仅用标准库，无需安装依赖。
"""
import csv, io, json, sys, os, collections, zipfile, re, glob, shutil, time, hashlib, traceback
import gzip as _gzip, base64 as _base64

# 被计划任务用 pythonw.exe（无控制台）拉起时 sys.stdout 为 None，会导致后续 print 崩溃；此处兜底
if getattr(sys, 'stdout', None) is None:
    try:
        sys.stdout = io.open(os.devnull, 'w', encoding='utf-8', errors='replace')
    except Exception:
        sys.stdout = io.StringIO()
if getattr(sys, 'stderr', None) is None:
    sys.stderr = sys.stdout
import datetime

BAD_DATES = collections.Counter()      # 形如『2026-10-30』这种不存在的日期（多为电脑系统日期不对导致）
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根目录（本脚本位于 脚本/ 子文件夹内）
HTML = os.path.join(HERE, "看板页面", "retail-dashboard.html")
DATA_JSON = os.path.join(HERE, "历史数据", "data.json")
# ★ 外置数据文件：登录后由用户上传到云端 Storage，看板运行态再拉取。
#   看板静态 HTML 不再内嵌任何业务数据/仓名/负责人，敏感信息全部经服务端鉴权后才下发。
DASHBOARD_JSON = os.path.join(HERE, "历史数据", "dashboard-data.json")
# ★ 历史数据仓库：看板所有历史数据都存在这里累积，删掉底表也不会丢
HISTORY_JSON = os.path.join(HERE, "历史数据", "历史数据存档.json")
BACKUP_DIR = os.path.join(HERE, "历史备份")
BACKUP_KEEP = 10          # 最多保留最近 10 份历史看板
# ★ 线上发布目录：数据更新后会自动把最新看板同步到这里，供线上发布使用
PUBLISH_DIR = os.path.join(HERE, "看板页面", "看板发布")
PUBLISH_HTML = os.path.join(PUBLISH_DIR, "index.html")
# ★ 月度目标配置：负责人 × 月的业绩目标（三平台营收之和），由人工维护，
#   更新时一并读入并内嵌进看板（供「负责人进度」页使用）。
TARGETS_PATH = os.path.join(HERE, "配置与映射", "月度目标配置.json")
# ★ 线上看板地址：更新完自动打开它（dm=1 会直接弹出「数据管理」提醒上传新数据）
ONLINE_URL = "https://l-rom0609.github.io/retail-monitor/"

# 底表文件夹：把新的订单底表 / 归属表直接放进这个文件夹即可（文件名任意）。
# 可用环境变量 RETAIL_DATA_DIR 覆盖（Mac 端指向你自己的底表目录，
# 例如 export RETAIL_DATA_DIR=~/retail-csv）；不设置时退回桌面「数据看板--底表更新」。
# 另外也支持仓库内的「底表输入」子文件夹（Mac 端把文件丢进仓库即可，无需固定桌面目录）。
DATA_DIR = os.environ.get("RETAIL_DATA_DIR") or os.path.join(os.path.expanduser("~"), "Desktop", "数据看板--底表更新")
LOCAL_INPUT = os.path.join(HERE, "数据底表")

# ==========================================================================
# ★★ 自配送=真实营收 的特例仓（2026-11-04 需求方确认，长期有效）★★
#    一般规则：AC 列配送平台 = 自配送 → 补单；
#    但下列仓的自配送单全部是【真实订单】，用户实付全部计入【真实营收】。
#    如需追加仓，照抄格式往集合里加（键 = 归属表 A 列「配送门店」名称）。
# ==========================================================================
SELF_DELIVERY_REAL_STORES = {"江苏省-南京市-江宁万达仓JM",
                             "广东省-揭阳市-商品城仓JM",   # 2026-11-06 需求方确认：自配送=真实（门店两个：普宁商品城 / 有人仓）
                             "广东省-普宁市-商品城仓JM",
                             "广东省-普宁市-商品城仓JM（有人仓）",
                             "广东省-汕头市-汕头仓JM"}

# ==========================================================================
# ★★ 永久店铺黑名单（需求方 2026-10-29 确认，长期有效）★★
#    以下【渠道店铺ID】（订单底表 A^ 列 = 第 56 列）的订单一律不计算进看板，
#    且以后每次自动更新也会持续剔除。
#    如需追加，照抄格式往集合里加即可，例如：{"26857582", "12345678"}
# ==========================================================================
EXCLUDE_STORE_IDS = {"26857582"}     # ← 被剥离的店铺：吉林省-长春市-诺瑞德仓JM 的其中一个门店
COL_STORE_ID = 55                    # A^「渠道店铺ID」所在列索引（0 开始）

# ==========================================================================
# ★★ 门店排除名单（2026-10-31 建档；2026-11-06 起已回归归属表的仓陆续移出）★★
#    这些配送门店没有对应的仓名（或属测试/无效数据），其订单一律不计入，
#    也不再出现在「未匹配」提醒里，避免每次更新都刷一屏噪音。
#    若以后要给其中某个门店建档：把它从本集合删掉，并在归属表里补一行即可。
# ==========================================================================
EXCLUDE_STORE_NAMES = {
    "0.96",
    "7.7",
    "AAA美丽大厦222",
    "云南省-昆明市-东旭骏城仓JM",
    "内蒙古-呼和浩特市-长乐宫仓JM",
    "少女情怀情趣成人用品(古城店·避孕套·飞机杯·跳蛋）",
    "山西省-太原市-中正天街仓JM",
    "广东省-佛山市-张槎仓JM",
    "广东省-佛山市-穗盐仓JM",
    "广东省-广州市-员岗仓JM",
    "广东省-广州市-车陂仓JM",
    "广东省-惠州市-三环天虹店仓JM",
    "广东省-梅州市-万达店",
    "广东省-深圳市-彩田仓JM",
    "广东省-深圳市-深圳东站仓JM",
    "广东省-深圳市-深大仓JM",
    "广东省-深圳市-科苑西仓JM",
    "广东省-深圳市-纷享城仓JM",
    "广东省-深圳市-龙塘仓JM",
    "广东省-珠海市-珠海店ST",
    "时尚家园仓",
    "江苏省-徐州市-云龙仓JM",
    "江苏省-盐城市-宝龙仓JM",
    "河北省-衡水市-桃城仓JM",
    "河南省-郑州市-郑东仓JM",
    "浙江省-杭州市-三坝仓JM",
    "浙江省-杭州市-西湖仓JM",
    "海南省-海口市-国贸仓JM",
    "福建省-厦门市-厦门站店JM",
    "福建省-泉州市-泰禾仓JM",
    "立基五金",
    "蕉互取货点（徐东仓）",
    "蕉互情趣取货点【万泰汇店】",
    "运城市-蕉互情趣尧梦湖店",
    "陕西省-西安市-大明宫仓JM",
    "陕西省-西安市-悦秀城仓JM",
    "陕西省-西安市-经开仓JM",
    "陕西省-西安市-长安仓JM",
}

# 兜底：文件夹不存在或为空时，回退到桌面上的原始两个文件
FALLBACK_CSV = os.path.join(os.path.expanduser("~"), "Desktop", "订单数据2026-10-28-d298ea.csv")
FALLBACK_XLSX = os.path.join(os.path.expanduser("~"), "Desktop", "仓名-负责人归属关系表.xlsx")


def pick_from_folder(folder, exts, keywords):
    """在底表文件夹里按扩展名找文件；命中多个时取「最后修改时间」最新的那个。
    keywords 用于优先挑选文件名里包含这些字的文件（如『订单』『归属』）。"""
    if not os.path.isdir(folder):
        return None
    found = [p for p in glob.glob(os.path.join(folder, '*'))
             if os.path.isfile(p)
             and os.path.splitext(p)[1].lower() in exts
             and not os.path.basename(p).startswith('~$')]     # 跳过 Excel 临时文件
    if not found:
        return None
    hit = [p for p in found if any(k in os.path.basename(p) for k in keywords)]
    pool = hit or found
    pool.sort(key=lambda p: os.path.getmtime(p))
    return pool[-1]


def list_csv_in_folder(folder):
    """列出底表文件夹里所有订单 csv（按修改时间从旧到新）。"""
    if not os.path.isdir(folder):
        return []
    found = [p for p in glob.glob(os.path.join(folder, '*'))
             if os.path.isfile(p)
             and os.path.splitext(p)[1].lower() == '.csv'
             and not os.path.basename(p).startswith('~$')]
    found.sort(key=lambda p: os.path.getmtime(p))
    return found


def norm_store_id(s):
    """规范店铺ID：去掉前后空格和 Excel 常见的文本前缀单引号。"""
    return (s or '').strip().lstrip("'").lstrip('\u2018\u2019\ufeff').strip()


def resolve_paths():
    """决定用哪些底表：订单表按 命令行指定 > 底表文件夹自动查找 > 桌面兜底；
    归属表【只认仓库内 配置与映射/仓名-负责人归属关系表.xlsx】——2026-11-06 用户规则：
    后续所有仓名-负责人归属关系一律以仓库该链路里的表为准，命令行/桌面副本一律忽略，
    避免两个版本互相覆盖（10-31~11-06 曾因桌面 217 行版与仓库 190 行版并存导致 30 仓历史被删）。
    订单表返回【列表】——文件夹里可以放多份 csv，会一次性全部读进来叠加。"""
    csv_list = []
    # 1) 命令行 / 拖拽传入（仅接受 csv；xlsx 参数已废弃，忽略并提示）
    for a in sys.argv[1:]:
        if a.startswith('-'):
            continue
        low = a.lower()
        if low.endswith('.csv'):
            csv_list.append(a)
        elif low.endswith('.xlsx') or low.endswith('.xls'):
            print("[提示] 归属表只认仓库内 配置与映射/仓名-负责人归属关系表.xlsx，已忽略参数：%s" % a)
    # 2) 底表文件夹自动查找（多份 csv 全收）
    if not csv_list:
        # 同时扫描桌面底表文件夹与仓库内「底表输入」子目录
        csv_list = list_csv_in_folder(DATA_DIR) + list_csv_in_folder(LOCAL_INPUT)
        # 文件名带「订单」的优先；若既有带订单的也有不带的，只用带订单的
        hit = [p for p in csv_list if '订单' in os.path.basename(p)]
        if hit and len(hit) != len(csv_list):
            csv_list = hit
    # 3) 订单表兜底
    if not csv_list and os.path.isfile(FALLBACK_CSV):
        csv_list = [FALLBACK_CSV]
    # 4) 归属表：唯一真源（必须存在，否则直接报错，不再悄悄退回桌面旧版）
    xlsx_p = os.path.join(HERE, "配置与映射", "仓名-负责人归属关系表.xlsx")
    if not os.path.isfile(xlsx_p):
        raise SystemExit(
            "[错误] 找不到唯一真源归属表：\n  %s\n"
            "请把《仓名-负责人归属关系表.xlsx》放入仓库 配置与映射/ 目录后重试。\n"
            "（桌面或其它位置的副本不会再被自动读取）" % xlsx_p)
    return csv_list, xlsx_p


CSV_PATH, XLSX_PATH = None, None    # 在 main() 里再解析，便于统一捕获异常
REPLACE_MODE = False                # --replace：清空历史重新累积
DROP_MONTHS = []                    # --drop 2026-09：删掉指定月份后重算
NO_OPEN = False                     # --no-open：后台自动更新时不弹浏览器


def read_id_map_from_html():
    """读取 仓ID → 仓名 的映射（迁移历史数据时保证仓名不丢）。
    看板已改为「数据外置」：仓名不再内嵌在 HTML 里，改为从 dashboard-data.json 读取。"""
    # 优先从外置数据文件读取
    if os.path.isfile(DASHBOARD_JSON):
        try:
            j = json.load(io.open(DASHBOARD_JSON, encoding='utf-8'))
            return {w['id']: w['name'] for w in j.get('WAREHOUSES', [])}
        except Exception:
            pass
    # 兜底：从旧版内嵌 HTML 读取（仅兼容历史遗留文件）
    try:
        html = io.open(HTML, encoding='utf-8').read()
        m = re.search(r'const WAREHOUSES = \[([\s\S]*?)\n  \];', html)
        if not m:
            return {}
        out = {}
        for line in m.group(1).split('\n'):
            mm = re.search(r'\{\s*"id"\s*:\s*"([^"]+)"\s*,\s*"name"\s*:\s*"([^"]*)"', line)
            if not mm:
                mm = re.search(r"\{\s*'id'\s*:\s*'([^']+)'\s*,\s*'name'\s*:\s*'([^']*)'", line)
            if mm:
                out[mm.group(1)] = mm.group(2)
        return out
    except Exception:
        return {}


def load_history():
    """读取历史数据仓库。没有的话，尝试从当前看板 HTML / data.json 迁移一份。"""
    if os.path.isfile(HISTORY_JSON):
        try:
            with io.open(HISTORY_JSON, encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, dict) and isinstance(d.get('records'), dict):
                return d, '数据仓库'
        except Exception as e:
            print('[提醒] 历史数据存档读取失败（%s），将尝试从当前看板迁移' % e)
    # 迁移：优先用 data.json
    src = None
    if os.path.isfile(DATA_JSON):
        try:
            with io.open(DATA_JSON, encoding='utf-8') as f:
                j = json.load(f)
            recs = j.get('records') or []
            whmap = {w['name']: w.get('manager', '') for w in (j.get('warehouses') or [])}
            # data.json 里没有 id→仓名 映射，按生成规则重建（warehouses 已是按仓名排序的）
            inv = {('W%03d' % (i + 1)): nm for i, nm in enumerate(sorted(whmap))}
            # 若能从看板 HTML 里读到真实映射，则以 HTML 为准（更可靠）
            hm = read_id_map_from_html()
            if hm:
                inv = hm
            store = {}
            for r in recs:
                nm = inv.get(r['w'], r['w'])
                key = '%s|%d|%s|%s' % (r['d'], r.get('h', 0), nm, r['p'])
                store[key] = {'o': r['o'], 'r': r['r'], 'ro': r['ro'], 'rr': r['rr'],
                              'so': r['so'], 'sr': r['sr']}
            src = {'records': store, 'warehouses': whmap,
                   'sources': [j.get('meta', {}).get('source', 'data.json')]}
            print('[迁移] 已从现有数据接管 %d 条记录作为历史存档起点' % len(store))
        except Exception as e:
            print('[提醒] 迁移 data.json 失败：%s' % e)
    if src is None:
        src = {'records': {}, 'warehouses': {}, 'sources': []}
    return src, '新建'


def merge_into_history(hist, new_agg, wh_mgr, csv_names):
    """把本次底表聚合结果合并进历史数据仓库。
    规则：以「日期 + 小时 + 仓名 + 平台」为唯一键，同键覆盖、异键追加。
    返回 (被覆盖的天数, 新增记录的日期集合)。"""
    recs = hist['records']
    whs = hist.setdefault('warehouses', {})
    covered_days = set()

    key_list = []
    for (d, wh, p, hh), v in new_agg.items():
        key = '%s|%d|%s|%s' % (d, hh, wh, p)
        existed = key in recs
        covered_days.add(d)
        recs[key] = {'o': v[0], 'r': round(v[1], 2),
                     'ro': v[0] - v[2], 'rr': round(v[1] - v[3], 2),
                     'so': v[2], 'sr': round(v[3], 2)}
        whs[wh] = wh_mgr.get(wh, whs.get(wh, ''))
        key_list.append((existed, d))
    # 记录来源文件
    src = hist.setdefault('sources', [])
    for n in csv_names:
        if n not in src:
            src.append(n)
    hist['sources'] = src[-20:]
    hist['updated'] = time.strftime('%Y-%m-%d %H:%M:%S')
    return len(covered_days), key_list


def drop_months(hist, prefixes):
    """按月份前缀（如 2026-09）删除历史数据。"""
    n = 0
    for k in list(hist['records'].keys()):
        if any(k.startswith(p) for p in prefixes):
            del hist['records'][k]
            n += 1
    return n


def history_to_records(hist):
    """把历史数据仓库还原成 records 列表 + 仓名清单。"""
    out = []
    for key, v in hist['records'].items():
        try:
            d, hh, wh, p = key.split('|', 3)
        except ValueError:
            continue
        out.append({'d': d, 'h': int(hh), 'w': wh, 'p': p,
                    'o': v['o'], 'r': v['r'], 'ro': v['ro'], 'rr': v['rr'],
                    'so': v['so'], 'sr': v['sr']})
    out.sort(key=lambda r: (r['d'], r['w'], r['p'], r['h']))
    return out


class Tee(object):
    """同时把输出写到屏幕和日志文件，窗口万一闪退也能留下证据。"""
    def __init__(self, path):
        self.console = sys.stdout
        try:
            self.f = io.open(path, 'w', encoding='utf-8')
        except Exception:
            self.f = None

    def write(self, s):
        try:
            self.console.write(s)
        except Exception:
            pass
        if self.f:
            try:
                self.f.write(s)
            except Exception:
                pass

    def flush(self):
        try:
            self.console.flush()
        except Exception:
            pass
        if self.f:
            try:
                self.f.flush()
            except Exception:
                pass

PLATFORMS = ["美团闪购", "淘宝闪购零售", "京东秒送"]
PLATFORM_COLORS = {"美团闪购": "#FFB020", "淘宝闪购零售": "#FF5A1F", "京东秒送": "#E1251B"}

# ==========================================================================
# ★ 平台卡片配色（需求方 2026-10-30 确认）：美团=黄底 / 淘宝=蓝底 / 京东=红底
#    这段会被写进看板的 PLATFORM_META，属于配置段，每次更新都重新生成——
#    改配色就改这里，否则会被下次更新覆盖回默认值。
# ==========================================================================
PLATFORM_STYLE = {
    "美团闪购": dict(
        color="#FFB020", accent="#F0A500", dot="#7A5200",
        bg="linear-gradient(135deg,#FFD24D 0%,#FFB020 100%)",
        line="rgba(120,80,0,.20)", name="#3D2A00", tagbg="rgba(255,255,255,.62)",
        ink="#2E1F00", muted="rgba(61,42,0,.74)", revbg="rgba(255,255,255,.60)",
        dot1="#FFFFFF", dot2="#8A5A00",
    ),
    "淘宝闪购零售": dict(
        color="#2E63D8", accent="#2E63D8", dot="#FFFFFF",
        bg="linear-gradient(135deg,#3A70E4 0%,#244CB8 100%)",
        line="rgba(255,255,255,.30)", name="#FFFFFF", tagbg="rgba(255,255,255,.22)",
        ink="#FFFFFF", muted="rgba(255,255,255,.86)", revbg="rgba(255,255,255,.16)",
        dot1="#FFFFFF", dot2="#A9C6FF",
    ),
    "京东秒送": dict(
        color="#E1251B", accent="#E2342A", dot="#FFFFFF",
        bg="linear-gradient(135deg,#F0503F 0%,#D01F16 100%)",
        line="rgba(255,255,255,.30)", name="#FFFFFF", tagbg="rgba(255,255,255,.22)",
        ink="#FFFFFF", muted="rgba(255,255,255,.86)", revbg="rgba(255,255,255,.16)",
        dot1="#FFFFFF", dot2="#FFBDB6",
    ),
}

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def read_xlsx_rows(path):
    """读取 xlsx 第一个 sheet，返回 [ {列字母: 值}, ... ]（含表头行）。"""
    z = zipfile.ZipFile(path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(NS + 'si'):
            shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    sheet_name = next(n for n in z.namelist() if re.match(r'xl/worksheets/sheet1\.xml$', n))
    sheet = ET.fromstring(z.read(sheet_name))
    rows = []
    for row in sheet.iter(NS + 'row'):
        cells = {}
        for c in row.findall(NS + 'c'):
            ref = c.get('r')
            ci = re.match(r'([A-Z]+)', ref).group(1)
            v = c.find(NS + 'v')
            t = c.get('t')
            if t == 'inlineStr':
                isn = c.find(NS + 'is')
                val = ''.join(x.text or '' for x in isn.iter(NS + 't')) if isn is not None else ''
            elif v is None:
                val = ''
            elif t == 's':
                val = shared[int(v.text)]
            else:
                val = v.text
            cells[ci] = val
        rows.append(cells)
    return rows


def norm_date(s):
    """J 列『2026/10/28 23:35』或『2026-10-28 23:35』→『2026-10-28』；异常返回 None。

    ★ 会校验日期是否真实存在（月 1~12、日 1~31）。
      电脑系统日期不正确时，导出的底表会带出『2026-10-30』这类非法日期，
      若不拦下来会污染数据仓库和看板的日期范围，必须丢弃并提示。
    """
    s = (s or '').strip()
    m = re.match(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s)
    if not m:
        return None
    y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mo <= 12 and 1 <= dd <= 31):
        return None
    try:
        import datetime as _dt
        _dt.date(y, mo, dd)          # 真实存在的日期（排除 2 月 30 日之类）
    except Exception:
        return None
    return '%04d-%02d-%02d' % (y, mo, dd)


def sha1_of(path):
    h = hashlib.sha1()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def backup_file(src_path, prefix):
    """通用备份：把 src_path 复制到备份目录（带时间戳），最多保留 BACKUP_KEEP 份。
    内容与最近一份完全相同时跳过。"""
    if not os.path.isfile(src_path):
        return None
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        olds = sorted(glob.glob(os.path.join(BACKUP_DIR, prefix + '_*')))
        if olds and sha1_of(olds[-1]) == sha1_of(src_path):
            return None                      # 和上一份一样，没必要重复备份
        dst = os.path.join(BACKUP_DIR, '%s_%s%s' % (
            prefix, time.strftime('%Y%m%d_%H%M%S'), os.path.splitext(src_path)[1]))
        shutil.copy2(src_path, dst)
        olds = sorted(glob.glob(os.path.join(BACKUP_DIR, prefix + '_*')))
        for p in olds[:-BACKUP_KEEP]:
            try:
                os.remove(p)
            except OSError:
                pass
        return dst
    except Exception as e:
        print('[提醒] 备份失败（不影响本次更新）：%s' % e)
        return None


def _md5(path):
    try:
        h = hashlib.md5()
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(1 << 20), b''):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def sync_publish_dir():
    """把最新看板同步到「看板发布/index.html」，作为线上发布的物料。

    每次数据更新都会自动执行：线上物料始终是最新的一版，
    发布动作由 WorkBuddy 完成（无需手工复制），网址固定不变。
    """
    try:
        if not os.path.isdir(PUBLISH_DIR):
            os.makedirs(PUBLISH_DIR)
        a, b = _md5(HTML), _md5(PUBLISH_HTML)
        if a and a == b:
            print('[发布同步] 线上物料已是最新，无需复制')
            return True
        shutil.copyfile(HTML, PUBLISH_HTML)
        size = os.path.getsize(PUBLISH_HTML) / 1048576.0
        print('[发布同步] 已更新线上物料：看板发布/index.html（%.1f MB）' % size)
        print('[发布同步] 提示：如需刷新线上分享链接，请在 WorkBuddy 里说「同步线上链接」')
        return True
    except Exception as e:
        print('[发布同步] 跳过（不影响本地看板）：%s' % e)
        return False


def _find_node():
    """找到本机可用的 node（WorkBuddy 自带优先，其次 PATH）"""
    try:
        import glob as _g
        pats = [os.path.join(os.path.expanduser('~'), '.workbuddy', 'binaries',
                             'node', 'versions', '*', 'node.exe')]
        for p in pats:
            for f in sorted(_g.glob(p), reverse=True):
                if os.path.isfile(f):
                    return f
    except Exception:
        pass
    try:
        import shutil as _s
        return _s.which('node') or _s.which('node.exe')
    except Exception:
        return None


def auto_sync_to_cloud():
    """把刚生成的 dashboard-data.json 自动上传到云端，线上看板即时生效。

    这样更新数据只需双击「一键更新看板」：算完直接上云，不必再手动拖文件。
    失败也不影响本地结果，可回退到看板里手动上传。
    """
    tool = os.path.join(HERE, '脚本', '自动同步', 'cloud_sync.js')
    if not os.path.isfile(tool):
        print('[云端同步] 未找到自动同步工具，已跳过（可在看板「数据管理」手动上传）')
        return False
    node = _find_node()
    if not node:
        print('[云端同步] 未找到 Node 运行环境，已跳过（可在看板「数据管理」手动上传）')
        return False
    try:
        import subprocess
        print('[云端同步] 正在把最新数据上传到云端…')
        flags = 0
        if os.name == 'nt':
            flags = 0x08000000      # CREATE_NO_WINDOW：不弹黑色控制台窗口
        r = subprocess.run([node, tool, '--once'],
                           cwd=os.path.join(HERE, '脚本', '自动同步'),
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           creationflags=flags, timeout=600)
        out = (r.stdout or b'').decode('utf-8', 'replace')
        for line in out.strip().splitlines():
            try:
                print('           ' + line)
            except Exception:
                pass
        if r.returncode == 0:
            print('[云端同步] ✔ 完成：线上看板已更新，同事刷新页面即可看到最新数据。')
            return True
        print('[云端同步] 未成功（原因见上）。可打开看板 →「数据管理」手动上传一次。')
        return False
    except Exception as e:
        print('[云端同步] 跳过（不影响本地数据）：%s' % e)
        return False


def embed_data_into_html_file(html_path, ext):
    """公开模式（2026-10-31 起）：把最新数据内嵌进看板 HTML 的注入点，替代云端外置。

    注入锚点：HTML 中以 `var __EMBEDDED_DATA__ = ...;` 至 `/*__DATA_END__*/` 包裹的整段，
    每次更新都会被本函数整体替换为最新数据。返回 True/False。
    """
    try:
        with io.open(html_path, encoding='utf-8') as f:
            html = f.read()
        raw = json.dumps(ext, ensure_ascii=False, separators=(',', ':'))
        # ★ 压缩内嵌：整包数据 gzip(level 9) + base64，首屏体积约为原来的 1/6，
        #   页面加载时用 DecompressionStream 解压（见 HTML 里的 gunzipText）。
        gz = _gzip.compress(raw.encode('utf-8'), 9)
        b64 = _base64.b64encode(gz).decode('ascii')
        block = ('var __EMBEDDED_GZ__ = "' + b64 + '";\n'
                 '  window.__PUBLIC_MODE__ = true;\n'
                 '  try { document.documentElement.classList.add("public-mode"); } catch(e){};\n'
                 '  // 数据源优先级：CDN → 同源文件 → 页面内嵌。\n'
                 '  // 原因：国内直连 GitHub Pages 实测仅 ~30KB/s（800KB 要 25s+）；\n'
                 '  //       走 jsDelivr 的 CDN 节点可达 200KB/s 以上，首屏快 5~8 倍。\n'
                 '  // 缓存用 no-cache（带 ETag 重验证）：数据没变只回 304，一变立刻拿到新版。\n'
                 '  var DATA_URLS = [\n'
                 '    "https://gcore.jsdelivr.net/gh/L-Rom0609/retail-monitor@main/dashboard-data.bin",\n'
                 '    "https://cdn.jsdelivr.net/gh/L-Rom0609/retail-monitor@main/dashboard-data.bin",\n'
                 '    "./dashboard-data.bin"\n'
                 '  ];\n'
                 '  function gunzipBytes(buf){\n'
                 '    return new Promise(function(resolve, reject){\n'
                 '      try{\n'
                 '        if (typeof DecompressionStream !== "function"){ reject(new Error("浏览器不支持解压，请升级到新版 Chrome / Edge / Safari")); return; }\n'
                 '        var u8 = new Uint8Array(buf || 0);\n'
                 '        if (u8.length < 2 || u8[0] !== 0x1f || u8[1] !== 0x8b){ reject(new Error("数据源返回了无效内容（非 gzip 数据）")); return; }\n'
                 '        var ds = new DecompressionStream("gzip"), w = ds.writable.getWriter();\n'
                 '        w.write(u8); w.close();\n'
                 '        new Response(ds.readable).text().then(resolve, reject);\n'
                 '      }catch(e){ reject(e); }\n'
                 '    });\n'
                 '  }\n'
                 '  async function fetchDataOne(url){\n'
                 '    var opt = { cache: "no-cache" };\n'
                 '    try { if (typeof AbortSignal !== "undefined" && AbortSignal.timeout) opt.signal = AbortSignal.timeout(12000); } catch(e){}\n'
                 '    var r = await fetch(url, opt);\n'
                 '    if (!r || !r.ok) throw new Error("HTTP " + (r && r.status));\n'
                 '    var ct = (r.headers.get("content-type") || ""), d;\n'
                 '    if (ct.indexOf("text") >= 0 || ct.indexOf("json") >= 0){\n'
                 '      var txt = await r.text();\n'
                 '      if (!txt || !txt.trim()) throw new Error("空数据");\n'
                 '      d = JSON.parse(await gunzipText(txt.trim()));\n'
                 '    } else {\n'
                 '      d = JSON.parse(await gunzipBytes(await r.arrayBuffer()));\n'
                 '    }\n'
                 '    if (!d || !d.DATA_DAYS) throw new Error("数据格式异常");\n'
                 '    window.__DATA_SRC__ = url;\n'
                 '    return d;\n'
                 '  }\n'
                 '  // 三个数据源（两个 CDN + 同源）全部并行竞速，谁先成功用谁；\n'
                 '  // 同源参与竞速可避免某个 CDN 请求挂起时整页永远卡在加载中\n'
                 '  function raceSources(urls){\n'
                 '    return new Promise(function(resolve, reject){\n'
                 '      var left = urls.length, errs = [];\n'
                 '      if (!left) return reject(new Error("无可用数据源"));\n'
                 '      urls.forEach(function(u){\n'
                 '        fetchDataOne(u).then(resolve, function(e){\n'
                 '          errs.push((u.split("/")[2] || u) + ": " + (e && e.message));\n'
                 '          if (--left <= 0) reject(new Error(errs.join(" | ")));\n'
                 '        });\n'
                 '      });\n'
                 '    });\n'
                 '  }\n'
                 '  async function loadDashboardData(){\n'
                 '    try { return await raceSources(DATA_URLS); }\n'
                 '    catch(e1){ console.warn("[看板] 三个数据源均不可达", e1 && e1.message); }\n'
                 '    if (typeof __EMBEDDED_GZ__ === "undefined" || !__EMBEDDED_GZ__) throw new Error("数据加载失败：CDN 与同源均不可达，且本页无内嵌备份");\n'
                 '    return JSON.parse(await gunzipText(__EMBEDDED_GZ__));\n'
                 '  }\n'
                 '  try { var __d = await loadDashboardData(); buildRecords(__d); }\n'
                 '  catch(e){ window.__EMBED_ERR__ = e; console.error("[看板] 数据加载失败", e);\n'
                 '    try { var _bm = document.getElementById("bootMask"); if (_bm) {\n'
                 '      var _m = _bm.querySelector(".bm"); if (_m) _m.textContent = "数据加载失败，请刷新页面重试";\n'
                 '      var _s = _bm.querySelector(".bt"); if (_s) _s.style.display = "none";\n'
                 '      var _t = document.createElement("div");\n'
                 '      _t.style.cssText = "margin-top:10px;color:#8FA6C6;font-size:12px;max-width:80vw;text-align:center";\n'
                 '      _t.textContent = String(e && e.message ? e.message : e);\n'
                 '      _bm.appendChild(_t);\n'
                 '    } } catch(_){}\n'
                 '  }\n'
                 '  /*__DATA_END__*/')
        key = 'var __EMBEDDED_GZ__ = '
        a = html.find(key)
        b = html.find('/*__DATA_END__*/', a) if a >= 0 else -1
        if a < 0 or b < 0:
            print('[内嵌] 未在 HTML 中找到注入锚点（__EMBEDDED_GZ__ / __DATA_END__），内嵌失败！')
            return False
        new_html = html[:a] + block + html[b + len('/*__DATA_END__*/'):]
        with io.open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_html)
        mb = os.path.getsize(html_path) / 1048576.0
        print('[内嵌] 已把最新数据写入 %s（%.2f MB，公开模式 · gzip 内嵌）' % (os.path.basename(html_path), mb))
        print('[内嵌] 压缩率：原始 %.2f MB → gzip %.2f MB（%.1f%%）'
              % (len(raw.encode('utf-8')) / 1048576.0, len(gz) / 1048576.0,
                 len(gz) * 100.0 / max(len(raw.encode('utf-8')), 1)))
        return True
    except Exception as e:
        print('[内嵌] 失败（看板数据未更新）：%s' % e)
        return False


def backup_html():
    """写入新数据前，先把当前这版看板备份一份（最多保留最近 BACKUP_KEEP 份）。"""
    return backup_file(HTML, 'retail-dashboard')


def backup_history():
    """历史数据仓库也备份一份（防止仓库文件损坏/误删）。"""
    return backup_file(HISTORY_JSON, '历史数据存档')


def load_targets():
    """读取「月度目标配置.json」：{ "2026-11": {"厚荣": 510000, ...}, ... }。

    该文件由人工维护（每个负责人每月一个「三平台营收之和」的目标值）。
    文件不存在或格式错误时返回空 dict，看板里对应格子显示「目标未设定」。
    """
    try:
        if os.path.isfile(TARGETS_PATH):
            with io.open(TARGETS_PATH, encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, dict):
                return d
            print('[提醒] 月度目标配置不是 JSON 对象，本版看板将不显示目标。')
    except Exception as e:
        print('[提醒] 月度目标配置读取失败（%s），本版看板将不显示目标。' % e)
    return {}


def norm_hour(s):
    s = (s or '').strip()
    m = re.search(r'(\d{1,2}):\d{2}', s)
    return int(m.group(1)) % 24 if m else 0


def main():
    global CSV_PATH, XLSX_PATH
    CSV_PATH, XLSX_PATH = resolve_paths()

    print('=' * 58)
    print(' 即时零售数据监控看板 —— 数据更新')
    print(' 运行时间：%s' % time.strftime('%Y-%m-%d %H:%M:%S'))
    print('=' * 58)

    if not CSV_PATH or not [p for p in CSV_PATH if os.path.isfile(p)]:
        raise SystemExit(
            '[没找到订单底表]\n'
            '请把导出的订单 CSV 放进这个文件夹：\n  %s\n'
            '（文件夹里至少要有一个 .csv 文件）\n'
            '提示：即使底表已经删除，看板里原有的历史数据也不会受影响。' % DATA_DIR)
    CSV_PATH = [p for p in CSV_PATH if os.path.isfile(p)]
    if not XLSX_PATH or not os.path.isfile(XLSX_PATH):
        raise SystemExit(
            '[没找到归属关系表]\n'
            '请把《仓名-负责人归属关系表.xlsx》放进仓库的这个目录：\n  %s\n'
            '（归属表唯一真源，桌面副本不再自动读取）' % os.path.join(HERE, '配置与映射'))

    if len(CSV_PATH) == 1:
        print('订单底表：%s' % os.path.basename(CSV_PATH[0]))
    else:
        print('订单底表：共 %d 份，本次全部读取并叠加' % len(CSV_PATH))
        for p in CSV_PATH:
            print('           - %s' % os.path.basename(p))
    print('归属关系：%s' % os.path.basename(XLSX_PATH))
    print('-' * 58)

    # ---------- 归属表 ----------
    mrows = read_xlsx_rows(XLSX_PATH)
    head = ' | '.join((mrows[0].get(c) or '').replace('\n', '') if mrows else '' for c in ('A', 'B', 'C'))
    if not head.strip(' |'):
        raise SystemExit('[归属表读不出来] 请确认表格第一行是表头，且 A/B/C 列分别为 配送门店 / 仓名 / 负责人')
    if '配送门店' not in head and '仓名' not in head:
        print('[提醒] 归属表表头是「%s」，和预期的 配送门店/仓名/负责人 不太一样，请确认列的顺序没错' % head)

    mapping = {}
    empty_mgr, no_wh = 0, 0
    for r in mrows[1:]:
        store = (r.get('A') or '').strip()
        wh = (r.get('B') or '').replace('\n', ' ').strip()
        mgr = (r.get('C') or '').strip()
        if not store:
            continue
        if not wh:
            no_wh += 1
            continue
        if not mgr:
            empty_mgr += 1
        mapping[store] = (wh, mgr)
    print('[归属表] %s ：表头「%s」' % (os.path.basename(XLSX_PATH), head))
    print('          有效映射 %d 条，仓名去重 %d 个' % (len(mapping), len(set(v[0] for v in mapping.values()))))
    if no_wh:
        print('          [提醒] 有 %d 行只有门店没有仓名，已跳过（这些门店的订单不会计入看板）' % no_wh)
    if empty_mgr:
        print('          [提醒] 有 %d 个仓没填负责人，看板里会显示为空白' % empty_mgr)
    if not mapping:
        raise SystemExit('[归属表是空的] 请检查 A 列（配送门店）和 B 列（仓名）是否有内容')

    # ---------- 订单底表 ----------
    def num(s):
        s = (s or '').replace(',', '').strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    def process_one_csv(cpath):
        """读取单个订单底表并完成聚合，返回 (agg, wh_mgr, 统计信息)。"""
        raw = open(cpath, 'rb').read()
        enc = None
        for e in ('utf-8-sig', 'gb18030'):
            try:
                text = raw.decode(e)
                enc = e
                break
            except UnicodeDecodeError:
                continue
        if enc is None:
            # 有些导出的表是「整份 UTF-8 + 少量 GBK 脏字节」，整文件解不出来，
            # 改为逐行容错：优先 UTF-8，失败的行单独用 GBK 兜底。
            lines, bad = [], 0
            for ln in raw.split(b'\n'):
                if not ln.strip():
                    continue
                try:
                    lines.append(ln.decode('utf-8-sig'))
                    continue
                except UnicodeDecodeError:
                    pass
                try:
                    lines.append(ln.decode('gb18030'))
                    continue
                except UnicodeDecodeError:
                    bad += 1
                    lines.append(ln.decode('utf-8-sig', errors='replace'))
            text = '\n'.join(lines)
            enc = 'utf-8-sig(逐行容错，%d 行含乱码已尽量救回)' % bad
        rows = list(csv.reader(io.StringIO(text)))
        hdr, data = rows[0], rows[1:]
        agg = collections.defaultdict(lambda: [0, 0.0, 0, 0.0])   # o, r, so, sr
        wh_mgr = {}
        skip_plat = skip_store = skip_bad = skip_excluded = skip_excluded_name = 0
        excluded_rev = 0.0
        unmatched = collections.Counter()
        for r in data:
            if len(r) < 29:
                skip_bad += 1
                continue
            p = r[1].strip()          # B 来源平台
            w = num(r[22])            # W 用户实付金额
            store = r[26].strip()     # AA 配送门店
            ac = r[28].strip()        # AC 配送平台
            d = norm_date(r[9])       # J 下单日期
            sid = (r[COL_STORE_ID].strip() if len(r) > COL_STORE_ID else '')
            if norm_store_id(sid) in EXCLUDE_STORE_IDS:
                skip_excluded += 1
                excluded_rev += w
                continue
            if p not in PLATFORMS:
                skip_plat += 1
                continue
            if store in EXCLUDE_STORE_NAMES:
                skip_excluded_name += 1
                continue
            if store not in mapping:
                skip_store += 1
                if store:
                    unmatched[store] += 1
                continue
            if d is None:
                skip_bad += 1
                _rd = (r[9] if len(r) > 9 else '').strip()
                if re.match(r'\d{4}[/-]\d{1,2}[/-]\d{1,2}', _rd):
                    BAD_DATES[_rd[:10].replace('/', '-')] += 1
                continue
            wh, mgr = mapping[store]
            wh_mgr[wh] = mgr
            a = agg[(d, wh, p, norm_hour(r[9]))]
            a[0] += 1
            a[1] += w
            # 自配送默认记补单；特例仓（SELF_DELIVERY_REAL_STORES）的自配送=真实订单
            if ac == '自配送' and store not in SELF_DELIVERY_REAL_STORES:
                a[2] += 1
                a[3] += w
        return agg, wh_mgr, {'rows': len(data), 'enc': enc, 'cols': len(hdr),
                             'skip_plat': skip_plat, 'skip_store': skip_store,
                             'skip_bad': skip_bad, 'skip_excluded': skip_excluded,
                             'skip_excluded_name': skip_excluded_name,
                             'excluded_rev': excluded_rev, 'unmatched': unmatched}

    # 逐个读取底表
    file_results = []
    for cpath in CSV_PATH:
        a_i, m_i, st = process_one_csv(cpath)
        file_results.append((cpath, a_i, m_i, st))
        print('[订单表] %s ：编码 %s，数据 %d 行，%d 列'
              % (os.path.basename(cpath), st['enc'], st['rows'], st['cols']))
        if st.get('skip_excluded'):
            print('  ↳ 按店铺黑名单剔除 %d 单（营收 ¥%.2f）｜黑名单：%s'
                  % (st['skip_excluded'], st['excluded_rev'], '、'.join(sorted(EXCLUDE_STORE_IDS))))

    # 多份底表按「日期」合并：后读取的表若覆盖同一天，则整体覆盖前面那天的旧数据
    agg = collections.defaultdict(lambda: [0, 0.0, 0, 0.0])
    wh_mgr = {}
    day_owner = {}
    replaced_days = []
    for cpath, a_i, m_i, st in file_results:
        days_i = sorted({k[0] for k in a_i})
        conflict = [d for d in days_i if d in day_owner]
        if conflict:
            replaced_days.append((os.path.basename(day_owner[conflict[0]]), os.path.basename(cpath), len(conflict)))
            for k in [k for k in agg if k[0] in conflict]:
                del agg[k]
        for k, v in a_i.items():
            a = agg[k]
            for i in range(4):
                a[i] += v[i]
        wh_mgr.update(m_i)
        for d in days_i:
            day_owner[d] = cpath

    tot_rows = sum(s['rows'] for _, _, _, s in file_results)
    skip_plat = sum(s['skip_plat'] for _, _, _, s in file_results)
    skip_store = sum(s['skip_store'] for _, _, _, s in file_results)
    skip_bad = sum(s['skip_bad'] for _, _, _, s in file_results)
    skip_excluded_name = sum(s.get('skip_excluded_name', 0) for _, _, _, s in file_results)
    unmatched_stores = collections.Counter()
    for _, _, _, s in file_results:
        unmatched_stores.update(s['unmatched'])

    if replaced_days:
        for old_f, new_f, n in replaced_days:
            print('[提醒] 「%s」与「%s」有 %d 个重复日期，已按后者为准（避免重复计入）' % (old_f, new_f, n))

    kept = tot_rows - skip_plat - skip_store - skip_bad - skip_excluded_name
    print('[口径] 剔除：非三平台 %d 单，门店未匹配 %d 单，格式异常 %d 单，名单内门店 %d 单；保留 %d 单'
          % (skip_plat, skip_store, skip_bad, skip_excluded_name, kept))
    if BAD_DATES:
        print('       [严重] 底表里出现 %d 单「不存在的日期」已丢弃，例如：%s'
              % (sum(BAD_DATES.values()), '、'.join('%s(%d 单)' % (k, v) for k, v in BAD_DATES.most_common(3))))
        print('              原因几乎都是「这台电脑的系统日期不正确」，导出平台按错误日期生成了底表。')
        print('              请校正电脑右下角的日期时间后重新导出底表，再更新一次看板。')
    if kept == 0:
        raise SystemExit('[没有可用数据] 所有订单都被剔除了，请检查底表里的门店名是否能和归属表对上')
    if unmatched_stores:
        print('       [提醒] 以下门店在归属表里查不到仓名（其订单未计入），'
              '如需计入请把它们补进归属表：')
        for s, c in unmatched_stores.most_common(8):
            print('         - %s（%d 单）' % (s, c))
        if len(unmatched_stores) > 8:
            print('         - ……其余 %d 个门店' % (len(unmatched_stores) - 8))

    T = [0, 0.0, 0, 0.0]
    for v in agg.values():
        for i in range(4):
            T[i] += v[i]
    new_days = sorted({k[0] for k in agg})
    print('[本次] 总营收 %.2f | 总单量 %d | 补单营收 %.2f | 真实营收 %.2f'
          % (T[1], T[0], T[3], T[1] - T[3]))
    print('[本次] 仓数 %d | 聚合记录 %d | 日期 %s ~ %s'
          % (len(wh_mgr), len(agg), new_days[0], new_days[-1]))

    # ---------- 合并进历史数据仓库 ----------
    hist, hist_from = load_history()
    if REPLACE_MODE:
        old_n = len(hist.get('records', {}))
        hist = {'records': {}, 'warehouses': {}, 'sources': []}
        print('[清空] 已丢弃原有历史数据 %d 条（--replace 模式）' % old_n)
    if DROP_MONTHS:
        n = drop_months(hist, DROP_MONTHS)
        print('[删除] 已移除 %s 的历史数据 %d 条' % ('、'.join(DROP_MONTHS), n))
    if not REPLACE_MODE:
        before = len(hist.get('records', {}))
        print('[仓库] 原有历史 %d 条（来源：%s）' % (before, hist_from))

    hist_days = {kk.split('|', 1)[0] for kk in hist.get('records', {})}
    overlapped = sum(1 for d in new_days if d in hist_days)
    merge_into_history(hist, agg, wh_mgr, [os.path.basename(p) for p in CSV_PATH])

    # ---------- 按当前归属表校核历史仓库 ----------
    # 需求：看板里只展示「归属表中存在对应仓名」的仓。
    # 所以在归属表里删掉的仓，它在历史月份里的数据也要一并移除。
    valid_wh = {v[0] for v in mapping.values()}
    stale = sorted({k.split('|')[2] for k in hist['records']} - valid_wh)
    if stale:
        stale_set = set(stale)
        rm_n = rm_o = 0
        rm_r = 0.0
        rm_days = set()
        for k in list(hist['records'].keys()):
            if k.split('|')[2] in stale_set:
                v = hist['records'].pop(k)
                rm_n += 1
                rm_o += v['o']
                rm_r += v['r']
                rm_days.add(k.split('|')[0])
        for w in stale:
            hist['warehouses'].pop(w, None)
        print('[校核] 以下仓在归属表里已不存在，其历史数据已同步移除：')
        for w in stale[:10]:
            print('         - %s' % w)
        if len(stale) > 10:
            print('         - ……其余 %d 个仓' % (len(stale) - 10))
        print('[校核] 共移除 %d 条记录（%.2f 元 / %d 单），涉及 %d 天（%s ~ %s）'
              % (rm_n, rm_r, rm_o, len(rm_days),
                 min(rm_days) if rm_days else '-', max(rm_days) if rm_days else '-'))

    _hist_bak = backup_history()          # 写入前先备份旧仓库（万一新数据有问题可回滚）
    with io.open(HISTORY_JSON, 'w', encoding='utf-8') as f:
        json.dump(hist, f, ensure_ascii=False, separators=(',', ':'))
    print('[仓库] 本次写入后共 %d 条（本次涉及 %d 天，其中 %d 天在库里已有、已覆盖更新）'
          % (len(hist['records']), len(new_days), overlapped))

    # ---------- 用历史全量数据导出 ----------
    all_records = history_to_records(hist)
    if not all_records:
        raise SystemExit('[没有可用数据] 数据仓库里没有记录，请检查底表和归属关系表')
    # 仓名清单：历史累积 ∪ 本次，负责人以最新归属表为准
    all_wh_mgr = dict(hist.get('warehouses', {}))
    all_wh_mgr.update(wh_mgr)
    whs = sorted(all_wh_mgr.items())
    name2id = {w[0]: 'W%03d' % (i + 1) for i, w in enumerate(whs)}

    records = []
    for x in all_records:
        records.append({
            'd': x['d'], 'h': x['h'], 'w': name2id.get(x['w'], 'W000'), 'p': x['p'],
            'o': x['o'], 'r': x['r'], 'ro': x['ro'], 'rr': x['rr'],
            'so': x['so'], 'sr': x['sr'],
        })

    # 全量汇总
    SUM = [0, 0.0, 0, 0.0]
    for x in records:
        SUM[0] += x['o']; SUM[1] += x['r']; SUM[2] += x['so']; SUM[3] += x['sr']
    all_days = sorted({x['d'] for x in records})
    print('[全量] 总营收 %.2f | 总单量 %d | 补单营收 %.2f | 真实营收 %.2f'
          % (SUM[1], SUM[0], SUM[3], SUM[1] - SUM[3]))
    print('[全量] 仓数 %d | 记录 %d | 日期 %s ~ %s（共 %d 天）'
          % (len(whs), len(records), all_days[0], all_days[-1], len(all_days)))
    months = sorted({d[:7] for d in all_days})
    print('[全量] 已存档月份：%s' % '、'.join(months))

    # 负责人按营收从高到低排序（下拉框里顺序更好看）；空名字排最后
    mgr_rev = collections.Counter()
    for x in all_records:                       # all_records 里是仓名，records 里已换成 ID
        mgr_rev[all_wh_mgr.get(x['w'], '')] += x['r']
    managers = [m for m, _ in sorted(mgr_rev.items(), key=lambda x: (-x[1], x[0]))]
    print('[负责人] ' + ' | '.join('%s %.0f元/%d仓' % (
        m or '(未填)', mgr_rev[m], sum(1 for w, mm in all_wh_mgr.items() if mm == m)) for m in managers))

    payload = {
        'meta': {
            'source': '、'.join(os.path.basename(p) for p in CSV_PATH),
            'mapping': os.path.basename(XLSX_PATH),
            'note': '仅三平台 + 仅归属表内仓名；全部订单状态计入；W列用户实付金额；自配送=补单'
                    '（特例：' + '、'.join(sorted(SELF_DELIVERY_REAL_STORES)) + ' 的自配送=真实）',
        },
        'platforms': PLATFORMS,
        'managers': managers,
        'warehouses': [{'name': k, 'manager': v} for k, v in whs],
        'records': records,
    }
    with io.open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, separators=(',', ':'))
    print('[导出] %s (%.1f KB)' % (DATA_JSON, os.path.getsize(DATA_JSON) / 1024))

        # ---------- 生成外置数据文件（dashboard-data.json）----------
    # 看板已改为「数据外置」：静态 HTML 不再内嵌任何业务数据 / 仓名 / 负责人，
    # 所有内容由用户登录后从云端 Storage 拉取。本步骤只产出外置 JSON，供用户上传。
    days = all_days
    # 紧凑编码：日期 / 仓 / 平台用索引，金额用「分」整数；运行态体积约为原始 JSON 的 1/5
    day_idx = {d: i for i, d in enumerate(days)}
    pf_idx = {p: i for i, p in enumerate(PLATFORMS)}
    wh_idx, wh_ids = {}, []
    crows = []
    for r in records:
        w = r['w']
        if w not in wh_idx:
            wh_idx[w] = len(wh_ids)
            wh_ids.append(w)
        crows.append([day_idx[r['d']], r['h'], wh_idx[w], pf_idx[r['p']],
                      int(round(r['rr'] * 100)), int(round(r['sr'] * 100)), r['ro'], r['so']])

    # 列式编码（v2）：把「行式」的 8 元组转成 8 条等长数组。
    # 同样的数据，gzip 后体积比行式小约 36%（同类数值在流里相邻，压缩率更高），
    # 前端 buildRecords 按下标直接读取，不需要重建二维表。
    if crows:
        _cols = [list(col) for col in zip(*crows)]
    else:
        _cols = [[] for _ in range(8)]

    # 数据源说明（随每次更新刷新）
    _srcs = '、'.join(os.path.basename(p) for p in CSV_PATH)
    _note = '数据源：%s（归属关系：%s；营收取 W 列「用户实付金额」，自配送=补单' \
            % (_srcs, os.path.basename(XLSX_PATH))
    if SELF_DELIVERY_REAL_STORES:
        _note += '；特例：%s 的自配送计入真实营收' % '、'.join(sorted(SELF_DELIVERY_REAL_STORES))
    _note += '；仅统计三平台与归属表内仓名'
    if EXCLUDE_STORE_IDS:
        _note += '；已剔除渠道店铺ID：%s' % '、'.join(sorted(EXCLUDE_STORE_IDS))
    _note += '）'
    if len(all_days) > 0:
        _note += '｜历史存档：%s ~ %s（%d 天）' % (all_days[0], all_days[-1], len(all_days))

    _mgr_whs = {}
    for _v in mapping.values():
        _mgr_whs.setdefault(_v[1], set()).add(_v[0])
    ext = {
        "DATA_MIN": all_days[0],
        "DATA_MAX": all_days[-1],
        "DATA_DAYS": all_days,
        "R_WH": wh_ids,
        "R_PF": PLATFORMS,
        # 只输出列式 C（体积比行式 R_ROWS 小 36%）；前端同时兼容两种格式，
        # 旧页面读不到本文件会回退内嵌数据，不会白屏。
        "C": _cols,
        "WAREHOUSES": [{"id": name2id[w], "name": w, "manager": m} for w, m in whs],
        # 2026-11-06 需求：归属表全量仓清单（含暂无数据的仓），hub「单仓360」与
        # 运营目标卡据此按归属表展示仓数，并对无数据仓在卡片中提醒
        "WAREHOUSES_ALL": [
            {"id": name2id.get(w, 'WX%03d' % (i + 1)), "name": w, "manager": _m}
            for i, (w, _m) in enumerate(sorted(
                {v[0]: v[1] for v in mapping.values()}.items()))
        ],
        "MANAGERS": managers,
        # 2026-11-06 需求：卡片直接展示「仓名-负责人归属关系表」里各负责人的仓数。
        # 2026-11-06 下午起归属表为门店维度（一仓可有多行门店），改为按 B 列仓名去重计数，
        # 与 WAREHOUSES_ALL、hub「单仓360」的仓名口径保持一致
        "MGR_WH_N": {m: len(ws) for m, ws in _mgr_whs.items()},
        "TARGETS": load_targets(),
        "SOURCE_NOTE": _note,
        # 口径覆盖率（P2-3）：本次更新处理了多少原始订单、剔了多少、为什么剔
        "COVERAGE": {
            "updated": time.strftime('%Y-%m-%d %H:%M'),
            "total": tot_rows,
            "kept": kept,
            "skip_plat": skip_plat,
            "skip_store": skip_store,
            "skip_bad": skip_bad,
            "skip_excluded_name": skip_excluded_name,
            "unmatched_stores": len(unmatched_stores),
            "unmatched_list": [s for s, _ in unmatched_stores.most_common(20)],
            "excluded_stores": len(EXCLUDE_STORE_NAMES),
            "mapping_rows": len(mapping),
        },
    }
    with io.open(DASHBOARD_JSON, 'w', encoding='utf-8') as f:
        json.dump(ext, f, ensure_ascii=False, separators=(',', ':'))
    saved = backup_html()
    print('[导出] %s (%.1f KB)' % (DASHBOARD_JSON, os.path.getsize(DASHBOARD_JSON) / 1024))
    print('-' * 58)
    if saved:
        print('[备份] 更新前的旧看板已留存：%s' % os.path.basename(saved))
    if _hist_bak:
        print('[备份] 更新前的历史数据仓库已留存：%s' % os.path.basename(_hist_bak))
    # ===== 公开模式（2026-10-31 起）：把最新数据内嵌进 HTML，不再依赖云端 =====
    embed_ok = embed_data_into_html_file(HTML, ext)
    print('[完成] 数据已更新！看板文件：%s' % HTML)
    sync_publish_dir()
    print('       历史数据合计：%s ~ %s（共 %d 天，%s）'
          % (days[0], days[-1], len(days), '、'.join(sorted({d[:7] for d in days}))))
    print('       总营收 %.2f 元 / 总单量 %d 单 / 仓数 %d 个'
          % (SUM[1], SUM[0], len(whs)))
    print('       看板文件：%s' % HTML)
    print('       数据仓库：%s  ← 看板的历史数据都存在这里，删底表也不会丢' % HISTORY_JSON)
    print('       历史备份：%s（最多留最近 %d 份，随时可双击打开回看旧数据）'
          % (BACKUP_DIR, BACKUP_KEEP))
    print('-' * 58)

    # 更新完自动用默认浏览器打开【线上看板】（失败也无妨，不影响更新结果）
    # 注意：不能打开本地 HTML（file:// 下云端登录会被 Origin 校验拒绝），必须走线上地址。
    # --no-open：后台自动更新时不弹浏览器（无人值守场景），看板自己会提示并刷新。
    if NO_OPEN:
        if embed_ok:
            print('（后台自动模式：不打开浏览器；最新数据已内嵌进看板并同步到发布目录，线上将自动重新发布）')
        else:
            print('（后台自动模式：内嵌失败，请检查上面的 [内嵌] 日志）')
        return
    try:
        import webbrowser
        webbrowser.open(ONLINE_URL)
        if embed_ok:
            print('已为您打开线上看板：最新数据已内嵌进看板，线上会自动重新发布。')
        else:
            print('已为您打开线上看板：但内嵌失败，请检查上面的 [内嵌] 日志。')
    except Exception:
        try:
            if os.name == 'nt':
                os.startfile(ONLINE_URL)
        except Exception:
            pass


if __name__ == '__main__':
    # 解析开关型参数
    REPLACE_MODE = any(a.lower() in ('--replace', '-r') for a in sys.argv[1:])
    NO_OPEN = any(a.lower() in ('--no-open', '--silent', '--auto') for a in sys.argv[1:]) \
              or os.environ.get('DASH_NO_OPEN') == '1'
    _i = 1
    while _i < len(sys.argv):
        if sys.argv[_i].lower() == '--drop' and _i + 1 < len(sys.argv):
            DROP_MONTHS += [x.strip() for x in sys.argv[_i + 1].split(',') if x.strip()]
            _i += 1
        _i += 1

    # 日志同时写一份到「更新日志.txt」（优先放在底表文件夹，方便查找）
    _log = os.path.join(DATA_DIR, '更新日志.txt') if os.path.isdir(DATA_DIR) \
        else os.path.join(HERE, '更新日志.txt')
    sys.stdout = Tee(_log)
    sys.stderr = sys.stdout

    try:
        main()
    except SystemExit as e:
        print('')
        print(e)
    except Exception as e:
        print('')
        print('[出错了] %s: %s' % (type(e).__name__, e))
        print('详细堆栈：')
        traceback.print_exc()
        print('请把「更新日志.txt」里的内容发给帮你做看板的人，方便定位问题。')

    print('')
    print('（本次完整记录已写入：%s）' % _log)

    if os.name == 'nt' and sys.stdin and sys.stdin.isatty():
        try:
            input('按回车键关闭窗口…')
        except Exception:
            pass
