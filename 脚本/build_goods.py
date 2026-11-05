# -*- coding: utf-8 -*-
"""商品页数据层：底表 → goods-data.bin
=====================================
口径（与页面一致，2026-11-05 由历史数据逐单校验锁定）：
  1. 一单一行：订单编号去重；只保留配送门店能在「仓名-负责人归属关系表」里找到的订单
  2. 用户实付 = 用户实付金额 × 100（分）；件数 = 数量 − 提示/赠品行数（名单见 配置与映射/赠品与提示行名单.txt）
  3. 配送平台=自配送 记为「补单」，其余为平台单
  4. 商品成本 = Σ(件数 × 单位成本)，成本取自 配置与映射/商品成本表.xlsx
  5. 仅「商品全部能识别位置码」的订单进入可核算实付/毛利
输出
  goods-data.bin（gzip JSON，供 goods.html 运行时拉取）

用法：python 脚本/build_goods.py
"""
import ast, collections, csv, glob, gzip, io, json, os, re, zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DESK = os.path.join(os.path.expanduser("~"), "Desktop")
DOWN = os.path.join(os.path.expanduser("~"), "Downloads")
csv.field_size_limit(20 * 1024 * 1024)   # 商品列可能很长，放宽单字段上限
XLSX_WH = os.path.join(HERE, "配置与映射", "仓名-负责人归属关系表.xlsx")
XLSX_COST = os.path.join(HERE, "配置与映射", "商品成本表.xlsx")
PROMPT_TXT = os.path.join(HERE, "配置与映射", "赠品与提示行名单.txt")
OUT_BIN = os.path.join(HERE, "goods-data.bin")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
TAIL = re.compile(r"(?<![A-Za-z0-9])([A-Za-z])[\s\-_/丨|]*(\d{1,3})([A-Za-z])?(-\d{1,3})?([\u4e00-\u9fa5]{0,8})\s*$")
_NEED = ("下单日期", "状态", "配送门店", "用户实付金额", "订单编号", "配送平台", "商品", "数量")


# ───────────── 通用读取 ─────────────
def read_text(p):
    b = open(p, "rb").read()
    # 平台导出的编码不统一：按月导出多为 UTF-8(BOM) 且带引号，按日导出多为 GBK；个别 UTF-16
    if b[:3] == b"\xef\xbb\xbf":
        return b.decode("utf-8-sig", "replace")
    if b[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return b.decode("utf-16", "replace")
    try:
        return b.decode("utf-8")
    except Exception:
        return b.decode("gb18030", "replace")


def read_xlsx(path):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    rows = []
    for row in ET.fromstring(z.read("xl/worksheets/sheet1.xml")).iter(NS + "row"):
        cells = {}
        for c in row.findall(NS + "c"):
            ref = c.get("r") or ""
            m = re.match(r"([A-Z]+)", ref)
            if not m:
                continue
            ci, t, v = m.group(1), c.get("t"), c.find(NS + "v")
            val = ""
            if t == "inlineStr":
                isn = c.find(NS + "is")
                val = "".join(x.text or "" for x in isn.iter(NS + "t")) if isn is not None else ""
            elif v is not None:
                if t == "s":
                    try:
                        val = shared[int(v.text)]
                    except Exception:
                        val = ""
                else:
                    val = v.text
            cells[ci] = (val or "").strip()
        rows.append(cells)
    return rows


def norm_date(v):
    s = (v or "").strip()[:10]
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3:
            return "%s-%02d-%02d" % (parts[0], int(parts[1]), int(parts[2]))
    return s


def _order_csvs():
    """收集全部订单底表：仓库底表目录 / 桌面归档目录 / 下载目录（重复单据靠订单号去重）。"""
    pats = [
        os.path.join(HERE, "数据底表", "**", "订单数据*.csv"),
        os.path.join(DESK, "数据看板--底表更新", "**", "订单数据*.csv"),
        os.path.join(DOWN, "订单数据*.csv"),
    ]
    seen, out = set(), []
    for pat in pats:
        for p in sorted(glob.glob(pat, recursive=True)):
            ap = os.path.abspath(p)
            if ap in seen:
                continue
            seen.add(ap)
            out.append((ap, "FIX1030" if "47-52" in os.path.basename(ap) else False))
    out.sort(key=lambda x: os.path.getmtime(x[0]))
    return out


# ───────────── 商品码解析 ─────────────
def norm_code(p):
    s = str(p or "").strip().upper()
    m = re.fullmatch(r"([A-Z]+)\s*-?\s*(\d{1,3})([A-Z])?\s*(-\s*\d{1,3})?", s)
    if m:
        out = m.group(1) + str(int(m.group(2))) + (m.group(3) or "")
        if m.group(4):
            out += "-" + str(int(m.group(4).lstrip("-").strip()))
        return out
    return None


def parse_code(head):
    m = TAIL.search(head)
    if not m:
        return None, None
    code = (m.group(1) + str(int(m.group(2))) + (m.group(3) or "") + (m.group(4) or "")).upper()
    name = (head[:m.start()] + m.group(5)).strip()
    name = re.sub(r"[丨|\s]+$", "", name).strip()
    return code, name


# ───────────── 配置：归属表 / 成本表 / 提示行 ─────────────
mapping = {}
for c in read_xlsx(XLSX_WH)[1:]:
    a, b, mgr = c.get("A") or "", (c.get("B") or "").replace("\n", " ").strip(), c.get("C") or ""
    if a and b:
        mapping[a] = (b, mgr)

# 与主看板完全一致的过滤口径：仅三平台 / 剔除黑名单店铺 / 自配送=补单（特例仓除外）
PLATFORMS = ("美团闪购", "淘宝闪购零售", "京东秒送")
_R = {}
for node in ast.parse(open(os.path.join(HERE, "脚本", "update_dashboard.py"), encoding="utf-8").read()).body:
    if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
        if node.targets[0].id in ("EXCLUDE_STORE_IDS", "EXCLUDE_STORE_NAMES", "SELF_DELIVERY_REAL_STORES"):
            _R[node.targets[0].id] = ast.literal_eval(node.value)
EXC_IDS = {str(x).strip() for x in _R.get("EXCLUDE_STORE_IDS", set())}
EXC_NAMES = {str(x).strip() for x in _R.get("EXCLUDE_STORE_NAMES", set())}
SELF_REAL = {str(x).strip() for x in _R.get("SELF_DELIVERY_REAL_STORES", set())}

PROMPTS = set()
if os.path.exists(PROMPT_TXT):
    for line in io.open(PROMPT_TXT, encoding="utf-8"):
        line = line.strip()
        if line and not line.startswith("#"):
            PROMPTS.add(line)


def is_prompt(head):
    return head in PROMPTS or ("提示" in head and head.startswith("【")) or head.startswith("【隐私配送】")


cost_map, cost_bad, cost_dup = {}, [], []
cost_rows = [c for c in read_xlsx(XLSX_COST)[1:] if any((c.get("A"), c.get("B"), c.get("C")))]
for c in cost_rows:
    enc, pos, cs = (c.get("A") or ""), (c.get("B") or ""), (c.get("C") or "")
    key = norm_code(pos)
    try:
        cent = round(float(cs) * 100)
    except Exception:
        cent = None
    if key is None or cent is None or cent < 0:
        cost_bad.append([enc, pos, cs])
        continue
    if key in cost_map:
        cost_dup.append([key, cost_map[key][0], cost_map[key][1], enc, pos, cs])
        continue
    cost_map[key] = [enc, cent]
MASTER = set(cost_map)

# ───────────── 逐单解析 ─────────────
FILES = [(p, fix) for p, fix in _order_csvs()]


def usable(path):
    try:
        hdr = next(csv.reader(io.StringIO(read_text(path))))
    except Exception:
        return False
    hs = {(h or "").strip() for h in hdr}
    return len(hdr) >= 50 and all(c in hs for c in _NEED)


kept = []
for p, fix in FILES:
    if fix == "FIX1030":
        # 文件名日期损坏的重复导出（2026-10-30），当天已有正常文件，整份跳过以免重复计数
        print("跳过（日期损坏的重复导出）：%s" % os.path.basename(p))
        continue
    if usable(p):
        kept.append((p, fix))
    else:
        print("跳过（非完整订单导出）：%s" % os.path.basename(p))
if not kept:
    raise SystemExit("未找到可用订单底表：请把「订单数据*.csv」放进 数据底表/ 或桌面「数据看板--底表更新」")
print("可用订单底表 %d 份" % len(kept))

recs = []
seen_oid = set()
skip = collections.Counter()
for path, fix in kept:
    rdr = csv.reader(io.StringIO(read_text(path)))
    hdr = next(rdr)
    h = {}
    for i, name in enumerate(hdr):
        name = (name or "").strip()
        if name and name not in h:
            h[name] = i
    i_d, i_st, i_amt = h["下单日期"], h["配送门店"], h["用户实付金额"]
    i_oid, i_pl, i_goods, i_qi, i_src = h["订单编号"], h["配送平台"], h["商品"], h["数量"], h["来源平台"]
    i_sid = h.get("渠道店铺ID")
    for r in rdr:
        need = max(i_d, i_st, i_amt, i_oid, i_pl, i_goods, i_qi, i_src)
        if len(r) <= need:
            continue
        if (r[i_src] or "").strip() not in PLATFORMS:
            skip["非三平台"] += 1
            continue
        if i_sid is not None and (r[i_sid] or "").strip().lstrip("'") in EXC_IDS:
            skip["黑名单店铺ID"] += 1
            continue
        store = (r[i_st] or "").strip()
        if store in EXC_NAMES:
            skip["黑名单店铺名"] += 1
            continue
        if store not in mapping:
            skip["门店不在归属表"] += 1
            continue
        oid = (r[i_oid] or "").strip().lstrip("'")
        d = "2026-10-30" if fix == "FIX1030" else norm_date(r[i_d])
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d or ""):
            skip["日期异常"] += 1
            continue
        key = (d, oid)
        if key in seen_oid:
            skip["重复单据"] += 1
            continue
        seen_oid.add(key)
        try:
            amt = round(float(r[i_amt] or 0) * 100)
        except Exception:
            amt = 0
        try:
            qi = int(float(r[i_qi] or 0))
        except Exception:
            qi = 0
        toks = []
        for seg in (r[i_goods] or "").split("+"):
            seg = seg.strip()
            if not seg:
                continue
            parts = seg.split("|")
            head = parts[0].strip()
            spec = parts[1].strip() if len(parts) > 1 else ""
            toks.append((head, spec, is_prompt(head)))
        is_bu = (r[i_pl] or "").strip() == "自配送" and store not in SELF_REAL
        recs.append((d, mapping[store][0], mapping[store][1], is_bu, amt, qi, toks))
print("逐单记录 %d 条 | 跳过 %s" % (len(recs), dict(skip) or "无"))
if not recs:
    raise SystemExit("没有解析到任何订单")

# ───────────── 聚合（口径同页面） ─────────────
freq = collections.Counter()
for d, wh, mgrname, is_bu, amt, qi, toks in recs:
    for head, spec, isp in toks:
        if isp:
            continue
        c, _ = parse_code(head)
        if c:
            freq[c] += 1

THREE_RE = re.compile(r"[A-Z]+\d{3}[A-Z]?(-\d{1,3})?")
TWO = re.compile(r"[A-Z]+\d{1,2}[A-Z]?(-\d{1,3})?")


def accepted(code):
    if not code:
        return False
    if code in MASTER:
        return True
    if TWO.fullmatch(code):
        return True
    return freq[code] >= 5


days, day_idx = [], {}
codes, code_idx = [], {}
wh_mgr = {}
cell = collections.defaultdict(lambda: [0, 0, 0, 0, 0, 0])
whd = collections.defaultdict(lambda: [0] * 8)
name_vote = collections.defaultdict(collections.Counter)
spec_vote = collections.defaultdict(collections.Counter)
seg_all = collections.Counter()
seg_full = collections.Counter()
scope = {"full": [0, 0], "part": [0, 0], "none": [0, 0]}
prompt_units = 0
three_new = collections.Counter()

for d, wh, mgrname, is_bu, amt, qi, toks in recs:
    if d not in day_idx:
        day_idx[d] = len(days)
        days.append(d)
    di = day_idx[d]
    wh_mgr[wh] = mgrname
    seg_all[d] += 1

    w = whd[(wh, di)]
    if is_bu:
        w[2] += 1
        w[3] += amt
    else:
        w[0] += 1
        w[1] += amt

    prompt_n = sum(1 for t in toks if t[2])
    prods = [t for t in toks if not t[2]]
    prompt_units += prompt_n
    if not prods:
        scope["none"][0] += 1
        scope["none"][1] += amt
        continue

    parsed = []
    for head, spec, _ in prods:
        c, name = parse_code(head)
        ok = accepted(c)
        if ok and c not in MASTER and THREE_RE.fullmatch(c):
            three_new[c] += 1
        parsed.append((c if ok else None, name, spec))
    hits = [p for p in parsed if p[0]]

    if not hits:
        scope["none"][0] += 1
        scope["none"][1] += amt
        continue
    if len(hits) < len(parsed):
        scope["part"][0] += 1
        scope["part"][1] += amt
        continue

    scope["full"][0] += 1
    scope["full"][1] += amt
    pool = max(0, qi - prompt_n)
    base, rem = divmod(pool, len(parsed))
    if is_bu:
        w[6] += 1
        w[7] += amt
    else:
        w[4] += 1
        w[5] += amt

    seen = set()
    for j, (code, name, spec) in enumerate(parsed):
        u = base + (1 if j < rem else 0)
        if code not in code_idx:
            code_idx[code] = len(codes)
            codes.append(code)
        ci = code_idx[code]
        cm = cost_map.get(norm_code(code))
        c_cent = (u * cm[1]) if cm else 0
        c_ = cell[(wh, ci, di)]
        if is_bu:
            c_[1] += u
            c_[3] += c_cent
            if code not in seen:
                c_[5] += 1
                seen.add(code)
        else:
            c_[0] += u
            c_[2] += c_cent
            if code not in seen:
                c_[4] += 1
                seen.add(code)
        if name:
            name_vote[code][name] += 1
        if spec:
            spec_vote[code][spec] += 1
    seg_full[d] += 1

days_old = days[:]
days = sorted(days_old)
order = {d: i for i, d in enumerate(days)}
cell = {(wh, ci, order[days_old[di]]): v for (wh, ci, di), v in cell.items()}
whd = {(wh, order[days_old[di]]): v for (wh, di), v in whd.items()}

seg = {t: {"orders": seg_all[t], "hit": seg_full[t]} for t in sorted(seg_all)}
prod = {}
for code, nv in name_vote.items():
    prod[code] = [nv.most_common(1)[0][0], spec_vote[code].most_common(1)[0][0] if spec_vote.get(code) else ""]
cost = [cost_map.get(norm_code(c)) for c in codes]

per_wh = collections.defaultdict(list)
for (wh, ci_, di_), v in cell.items():
    per_wh[wh].append([ci_, di_] + v)
whs = []
for wh in sorted(per_wh, key=lambda w: -sum(x[2] + x[3] for x in per_wh[w])):
    flat = []
    for x in sorted(per_wh[wh]):
        flat.extend(x)
    wdays = [[di_] + v for (w2, di_), v in sorted(whd.items(), key=lambda kv: kv[0][1]) if w2 == wh]
    whs.append([wh, wh_mgr.get(wh, ""), flat, wdays])

all_units = sum(x[2] + x[3] for v in per_wh.values() for x in v)
all_lines = sum(x[6] + x[7] for v in per_wh.values() for x in v)
all_cost = sum(x[4] + x[5] for v in per_wh.values() for x in v) / 100
all_hit_amt = sum(v[5] + v[7] for v in whd.values()) / 100
all_amt = sum(v[1] + v[3] for v in whd.values()) / 100
all_hit_orders = sum(v[4] + v[6] for v in whd.values())

payload = {
    "meta": {
        "range": days[0] + " – " + days[-1],
        "days": days, "codes": codes, "cost": cost,
        "costStat": {"rows": len(cost_rows), "matched": sum(1 for x in cost if x),
                     "missing": sum(1 for x in cost if not x), "bad": cost_bad, "dup": cost_dup},
        "segs": seg, "orders": len(recs), "hit": sum(seg_full.values()),
        "wh": len(whs), "nCodes": len(codes),
        "lines": all_lines, "units": int(all_units), "promptUnits": prompt_units,
        "scope": {"full": scope["full"], "part": scope["part"], "none": scope["none"]},
        "allOrders": sum(seg_all.values()),
        "amt": round(all_hit_amt), "amtAll": round(all_amt), "costTotal": round(all_cost),
        "hitOrders": all_hit_orders,
    },
    "prod": prod,
    "wh": whs,
}
raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
gz = gzip.compress(raw, 9)
open(OUT_BIN, "wb").write(gz)

m = payload["meta"]
print("商品页数据：%s ~ %s（%d 天）| 仓 %d | 商品 %d 个" % (days[0], days[-1], len(days), len(whs), len(codes)))
print("订单 %d（可核算 %d）| 实付 ¥%s / 全部 ¥%s | 成本 ¥%s | 件 %s | 行 %s" % (
    m["orders"], m["hit"], format(m["amt"], ","), format(m["amtAll"], ","),
    format(m["costTotal"], ","), format(m["units"], ","), format(m["lines"], ",")))
print("成本表 有效 %d / 缺成本 %d / 无法归一 %d / 重复 %d | 提示行 %d 行" % (
    len(cost_map), m["costStat"]["missing"], len(cost_bad), len(cost_dup), prompt_units))
print("已写出 %s（JSON %.1f MB → gzip %.1f MB）" % (
    os.path.basename(OUT_BIN), len(raw) / 1048576, len(gz) / 1048576))
