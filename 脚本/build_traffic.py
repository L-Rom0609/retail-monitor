# -*- coding: utf-8 -*-
"""美团闪购「仓流量」数据层（仓库版）
=====================================
关联链：门店明细 D列「商家ID」 == 订单底表「渠道店铺ID」 -> 「配送门店」 -> 归属表 -> 仓名/负责人

输入
  流量：数据底表/门店明细/*.csv（平台导出，每月一份，文件名 YYYYMMDD-YYYYMMDD-全部门店-门店明细-*.csv）
  订单：数据底表/**.csv → 桌面「数据看板--底表更新」→ 下载目录（自动发现；同一 日期+订单号 会去重）

输出
  历史数据/traffic_payload.json —— 供 脚本/build_traffic_page.py 生成线上页 traffic.html

用法：python 脚本/build_traffic.py
"""
import csv, io, json, collections, re, zipfile, os, glob
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DESK = os.path.join(os.path.expanduser("~"), "Desktop")
DOWN = os.path.join(os.path.expanduser("~"), "Downloads")
TRAFFIC_DIR = os.path.join(HERE, "数据底表", "门店明细")


def _traffic_csvs():
    """收集全部「全部门店」流量明细。

    平台按月/按日导出，处理完会被归档到桌面「数据看板--底表更新」或留在下载目录，
    只认「数据底表/门店明细」会导致历史月份丢失（表现为流量页只剩最新一天）。
    这里按 (起, 止) 去重，并跳过与已选区间重叠的文件（避免同一天被重复累加）。
    """
    pats = [
        os.path.join(TRAFFIC_DIR, "*.csv"),
        os.path.join(DESK, "数据看板--底表更新", "**", "*门店明细*.csv"),
        os.path.join(DOWN, "*门店明细*.csv"),
    ]
    rx = re.compile(r"(\d{8})-(\d{8})-全部门店-门店明细")
    best = {}
    for pat in pats:
        for p in glob.glob(pat, recursive=True):
            m = rx.search(os.path.basename(p))
            if not m:
                continue
            key = (m.group(1), m.group(2))
            try:
                score = (1 if os.path.abspath(p).startswith(HERE) else 0, os.path.getsize(p))
            except OSError:
                continue
            if key not in best or score > best[key][0]:
                best[key] = (score, p)
    items = sorted((k[0], k[1], v[1]) for k, v in best.items())
    picked, spans = [], []
    for s, e, p in items:
        if any(not (e < s2 or s > e2) for s2, e2 in spans):
            print("跳过（区间重叠）：%s" % os.path.basename(p))
            continue
        picked.append(p)
        spans.append((s, e))
    return picked


TRAFFIC_FILES = _traffic_csvs()
XLSX = os.path.join(HERE, "配置与映射", "仓名-负责人归属关系表.xlsx")
OUT_JSON = os.path.join(HERE, "历史数据", "traffic_payload.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _order_csvs():
    """自动收集全部订单底表。重复文件不会重复计数（扫描时按 日期+订单号 去重）。"""
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


if not TRAFFIC_FILES:
    raise SystemExit("未找到流量明细：请把「门店明细」CSV 放进 数据底表/门店明细/")
FILES = _order_csvs()
if not FILES:
    raise SystemExit("未找到订单底表：请把「订单数据*.csv」放进 数据底表/ 或桌面「数据看板--底表更新」")
print("流量明细 %d 份 | 订单底表 %d 份" % (len(TRAFFIC_FILES), len(FILES)))


def read_text(p):
    raw = open(p, "rb").read()
    if raw[:3] == b"\xef\xbb\xbf":
        return raw.decode("utf-8-sig", errors="replace")
    try:
        return raw.decode("gb18030")
    except Exception:
        return raw.decode("utf-8", errors="replace")


_NEED = ("下单日期", "来源平台", "配送门店", "用户实付金额", "订单编号", "配送平台", "渠道店铺ID")


def _usable_order_csv(path):
    try:
        hdr = next(csv.reader(io.StringIO(read_text(path))))
    except Exception:
        return False
    hs = {(h or "").strip() for h in hdr}
    return len(hdr) >= 50 and all(c in hs for c in _NEED)


_kept = []
for _p, _fix in FILES:
    if _usable_order_csv(_p):
        _kept.append((_p, _fix))
    else:
        print("跳过（非完整订单导出，列数不符）：%s" % os.path.basename(_p))
FILES = _kept
if not FILES:
    raise SystemExit("没有可用的订单底表（需要平台完整导出，58 列）")
print("可用订单底表 %d 份" % len(FILES))


def norm_date(v):
    """底表日期格式不统一：2026-11-02 / 2026/11/2 等，统一为 ISO。"""
    s = (v or "").strip()[:10]
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3:
            return "%s-%02d-%02d" % (parts[0], int(parts[1]), int(parts[2]))
    return s


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
            ci = re.match(r"([A-Z]+)", c.get("r")).group(1)
            v = c.find(NS + "v"); t = c.get("t"); val = ""
            if t == "inlineStr":
                isn = c.find(NS + "is")
                val = "".join(x.text or "" for x in isn.iter(NS + "t")) if isn is not None else ""
            elif v is not None:
                val = shared[int(v.text)] if t == "s" else v.text
            cells[ci] = (val or "").strip()
        rows.append(cells)
    return rows


# ── 归属表 ──
mapping = {}
for c in read_xlsx(XLSX)[1:]:
    a, b, m = c.get("A") or "", (c.get("B") or "").replace("\n", " ").strip(), c.get("C") or ""
    if a and b:
        mapping[a] = (b, m)

# ── 排除口径（与主看板 update_dashboard.py 完全一致） ──
import ast
src_py = open(os.path.join(HERE, "脚本", "update_dashboard.py"), encoding="utf-8").read()
_R = {}
for node in ast.parse(src_py).body:
    if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
        if node.targets[0].id in ("EXCLUDE_STORE_IDS", "EXCLUDE_STORE_NAMES", "SELF_DELIVERY_REAL_STORES"):
            _R[node.targets[0].id] = ast.literal_eval(node.value)
EXC_IDS = {str(x) for x in _R.get("EXCLUDE_STORE_IDS", set())}
EXC_NAMES = set(_R.get("EXCLUDE_STORE_NAMES", []))
SELF_REAL = set(_R.get("SELF_DELIVERY_REAL_STORES", []))

all_dates = set()
for p in TRAFFIC_FILES:
    t = read_text(p)
    rdr = csv.reader(io.StringIO(t))
    next(rdr)
    for r in rdr:
        d = (r[0] or "").strip()
        if len(d) == 8 and d.isdigit():
            all_dates.add(d[:4] + "-" + d[4:6] + "-" + d[6:8])
DAYS = sorted(all_dates)
DIDX = {d: i for i, d in enumerate(DAYS)}
DIDSET = set(DAYS)
ND = len(DAYS)
print("流量明细覆盖 %s ~ %s，共 %d 天" % (DAYS[0], DAYS[-1], ND))
# ── sid -> 门店（全期底表） ──
sid_store = {}
for path, fix in FILES:
    rdr = csv.reader(io.StringIO(read_text(path)))
    hdr = next(rdr)
    ix = {}
    for i, h in enumerate(hdr):
        h = h.strip()
        if h and h not in ix:
            ix[h] = i
    i_sid, i_st = ix.get("渠道店铺ID"), ix["配送门店"]
    if i_sid is None:
        continue
    for r in rdr:
        if len(r) <= max(i_sid, i_st):
            continue
        sid = (r[i_sid] or "").strip().lstrip("'")
        if sid:
            sid_store.setdefault(sid, (r[i_st] or "").strip())

# ── 全期美团营收/单量（底表口径：逐单扫全部底表，仅美团闪购） ──
order_by_wh = collections.defaultdict(lambda: [0, 0, 0, 0])     # 单量, 营收分, 真实单, 真实营收分
order_by_wh_d = collections.defaultdict(lambda: [0, 0])         # (仓, ISO日期) -> [单量, 营收分]
seen_oid = set()
for path, fix in FILES:
    rdr = csv.reader(io.StringIO(read_text(path)))
    hdr = next(rdr)
    ix = {}
    for i, h in enumerate(hdr):
        h = h.strip()
        if h and h not in ix:
            ix[h] = i
    i_d, i_amt, i_st, i_src, i_oid = ix["下单日期"], ix["用户实付金额"], ix["配送门店"], ix["来源平台"], ix["订单编号"]
    i_sid, i_pl = ix.get("渠道店铺ID"), ix["配送平台"]
    for r in rdr:
        if len(r) <= max(i_d, i_amt, i_st, i_src, i_oid):
            continue
        if (r[i_src] or "").strip() != "美团闪购":
            continue
        if i_sid is not None and (r[i_sid] or "").strip().lstrip("'") in EXC_IDS:
            continue
        store = (r[i_st] or "").strip()
        if store in EXC_NAMES or store not in mapping:
            continue
        d = "2026-10-30" if fix == "FIX1030" else norm_date(r[i_d])
        if d not in DIDSET:
            continue
        oid = (r[i_oid] or "").strip().lstrip("'")
        if oid and (d, oid) in seen_oid:
            continue
        if oid:
            seen_oid.add((d, oid))
        try:
            amt = round(float((r[i_amt] or "0").replace(",", "")) * 100)
        except Exception:
            amt = 0
        wh = mapping[store][0]
        v = order_by_wh[wh]
        v[0] += 1; v[1] += amt
        vd = order_by_wh_d[(wh, d)]
        vd[0] += 1; vd[1] += amt
        if (r[i_pl] or "").strip() != "自配送" or store in SELF_REAL:
            v[2] += 1; v[3] += amt

# ── 门店明细：5 个月份文件，按商家×日期聚合 ──
F = dict(expo=13, visit=14, buyers=15, new=16, orders=19, price=20, pay=21, sub=22,
         c_cancel=23, u_cancel=24, late=25, oos=26, wrong=27, quality=28, d_cancel=29,
         wait=30, deliver=31, im=32, im_r=33, im_1m=35, bad=36, comp=37,
         s_score=38, g_score=39, d_score=40, dur=41)



def num(v):
    try:
        return float((v or "0").replace(",", "").replace("%", ""))
    except Exception:
        return 0.0


stores = collections.defaultdict(lambda: {
    "name": "", "sid": "", "city": "", "new": 0, "days": 0, "online_h": 0.0, "plan_h": 0.0,
    "expo": 0, "visit": 0, "buyers": 0, "newb": 0, "orders": 0, "price": 0.0, "pay": 0.0,
    "sub": 0.0, "c_cancel": 0, "u_cancel": 0, "late": 0, "oos": 0, "wrong": 0, "quality": 0,
    "d_cancel": 0, "wait": 0.0, "deliver": 0.0, "im": 0, "im_r": 0, "im_1m": 0,
    "bad": 0, "comp": 0, "s_score": 0.0, "g_score": 0.0, "d_score": 0.0, "dur": 0.0,
})
daily = collections.defaultdict(lambda: [0, 0, 0, 0, 0])   # expo, visit, buyers, orders, pay
# 门店级按日（19 字段同主口径）
sid_daily = collections.defaultdict(dict)

for p in TRAFFIC_FILES:
    rdr = csv.reader(io.StringIO(read_text(p)))
    next(rdr)
    for r in rdr:
        d = (r[0] or "").strip()
        if len(d) != 8 or not d.isdigit():
            continue
        di = DIDX.get(d[:4] + "-" + d[4:6] + "-" + d[6:8])
        if di is None:
            continue
        sid = r[3]
        s = stores[sid]
        if not s["name"]:
            s["name"] = r[2]; s["sid"] = sid
            s["city"] = (r[4] or "") + "·" + (r[5] or "")
            s["first"] = r[7]
        s["new"] = 1 if r[6] == "是" else s["new"]
        s["days"] += 1
        s["online_h"] += num(r[11]); s["plan_h"] += num(r[12])
        for k in ("expo", "visit", "buyers", "newb", "orders", "c_cancel", "u_cancel", "late",
                  "oos", "wrong", "quality", "d_cancel", "im", "im_r", "im_1m", "bad", "comp"):
            src = {"newb": "new"}.get(k, k)
            s[k] += int(num(r[F[src]]))
        for k in ("price", "pay", "sub", "wait", "deliver", "dur"):
            s[k] += num(r[F[k]])
        od = int(num(r[19]))
        s["s_score"] += num(r[38]) * od
        s["g_score"] += num(r[39]) * od
        s["d_score"] += num(r[40]) * od
        dd = daily[di]
        dd[0] += int(num(r[13])); dd[1] += int(num(r[14])); dd[2] += int(num(r[15]))
        dd[3] += int(num(r[19])); dd[4] += num(r[21])
        dv = sid_daily[sid].get(di)
        if dv is None:
            dv = sid_daily[sid][di] = [0.0] * 19
        dv[0] += int(num(r[13])); dv[1] += int(num(r[14])); dv[2] += int(num(r[15])); dv[3] += int(num(r[16]))
        dv[4] += int(num(r[19])); dv[5] += num(r[21]); dv[6] += num(r[20]); dv[7] += num(r[22])
        dv[8] += int(num(r[23])); dv[9] += int(num(r[25])); dv[10] += int(num(r[26]))
        dv[11] += int(num(r[36])); dv[12] += int(num(r[37])); dv[13] += int(num(r[32])); dv[14] += int(num(r[35]))
        dv[15] += num(r[11]); dv[16] += num(r[12]); dv[17] += num(r[38]) * od; dv[18] += num(r[41])

# 门店按日 -> 仓按日（稀疏）
wh_daily = collections.defaultdict(dict)
for sid, daymap in sid_daily.items():
    st = sid_store.get(sid)
    wh = mapping.get(st, (None, None))[0] if st else None
    if not wh:
        continue
    tgt = wh_daily[wh]
    for di, dv in daymap.items():
        cur = tgt.get(di)
        if cur is None:
            tgt[di] = list(dv)
        else:
            for k in range(19):
                cur[k] += dv[k]

# 底表按日 -> {仓: {天序号: [单量, 营收元]}}
ordd = collections.defaultdict(dict)
for (wh, d), v in order_by_wh_d.items():
    ordd[wh][DIDX[d]] = [v[0], round(v[1] / 100)]

# ── 商家 -> 仓（全期聚合） ──
whs = collections.defaultdict(lambda: {
    "stores": 0, "expo": 0, "visit": 0, "buyers": 0, "newb": 0, "porders": 0,
    "price": 0.0, "pay": 0.0, "sub": 0.0, "c_cancel": 0, "u_cancel": 0, "late": 0,
    "oos": 0, "wrong": 0, "quality": 0, "d_cancel": 0, "im": 0, "im_r": 0, "im_1m": 0,
    "bad": 0, "comp": 0, "s_score": 0.0, "g_score": 0.0, "d_score": 0.0,
    "wait": 0.0, "deliver": 0.0, "dur": 0.0, "days": 0, "src": set(),
    "online_h": 0.0, "plan_h": 0.0,
})
unmatch = collections.defaultdict(float)
unmatch["names"] = []
for sid, s in stores.items():
    st = sid_store.get(sid)
    wh = mapping.get(st, (None, None))[0] if st else None
    target = whs[wh] if wh else unmatch
    if wh:
        whs[wh]["stores"] += 1
        whs[wh]["src"].add(sid)
    else:
        unmatch["stores"] += 1
        if len(unmatch["names"]) < 8:
            unmatch["names"].append(s["name"][:34])
    for k in ("expo", "visit", "buyers", "newb", "porders", "price", "pay", "sub", "c_cancel",
              "u_cancel", "late", "oos", "wrong", "quality", "d_cancel", "im", "im_r", "im_1m",
              "bad", "comp", "s_score", "g_score", "d_score", "wait", "deliver", "dur", "days",
              "online_h", "plan_h"):
        target[k] += s["orders"] if k == "porders" else s[k]

out_wh = []
for wh, v in whs.items():
    od_base = order_by_wh.get(wh, [0, 0, 0, 0])
    o = v["porders"]
    out_wh.append({
        "name": wh, "mgr": mapping.get(wh, ("", ""))[1] if wh in mapping else "",
        "stores": v["stores"],
        "expo": v["expo"], "visit": v["visit"], "buyers": v["buyers"], "newb": v["newb"],
        "porders": o,
        "rev": round(od_base[1] / 100), "ord": od_base[0],   # 底表口径（美团闪购）
        "real_ord": od_base[2], "real_rev": round(od_base[3] / 100),
        "pay": round(v["pay"]), "price": round(v["price"]), "sub": round(v["sub"]),
        "c_cancel": v["c_cancel"], "u_cancel": v["u_cancel"], "late": v["late"],
        "oos": v["oos"], "wrong": v["wrong"], "quality": v["quality"], "d_cancel": v["d_cancel"],
        "im": v["im"], "im_r": v["im_r"], "im_1m": v["im_1m"],
        "bad": v["bad"], "comp": v["comp"],
        "s_score": round(v["s_score"] / o, 2) if o else 0,
        "g_score": round(v["g_score"] / o, 2) if o else 0,
        "d_score": round(v["d_score"] / o, 2) if o else 0,
        "wait": round(v["wait"] / v["days"], 1) if v["days"] else 0,
        "deliver": round(v["deliver"] / v["days"], 1) if v["days"] else 0,
        "dur": round(v["dur"] / v["days"], 1) if v["days"] else 0,
        "avg_online": round(v["online_h"] / v["days"], 1) if v["days"] else 0,
        "avg_plan": round(v["plan_h"] / v["days"], 1) if v["days"] else 0,
    })

# 底表有美团单但流量表里没有的仓（多为商家账号不在导出范围）
_covered = {w["name"] for w in out_wh}
for wh, od_base in order_by_wh.items():
    if wh in _covered:
        continue
    out_wh.append({
        "name": wh, "mgr": mapping.get(wh, ("", ""))[1],
        "stores": 0, "expo": 0, "visit": 0, "buyers": 0, "newb": 0, "porders": 0,
        "rev": round(od_base[1] / 100), "ord": od_base[0],
        "real_ord": od_base[2], "real_rev": round(od_base[3] / 100),
        "pay": 0, "price": 0, "sub": 0, "c_cancel": 0, "u_cancel": 0, "late": 0,
        "oos": 0, "wrong": 0, "quality": 0, "d_cancel": 0, "im": 0, "im_1m": 0,
        "bad": 0, "comp": 0, "s_score": 0, "g_score": 0, "d_score": 0,
        "wait": 0, "deliver": 0, "dur": 0, "avg_online": 0, "avg_plan": 0, "no_traffic": 1,
    })

# 门店级输出（含未匹配，标记 wh）
out_store = []
for sid, s in stores.items():
    st = sid_store.get(sid)
    wh = mapping.get(st, (None, None))[0] if st else None
    o = s["orders"]
    out_store.append({
        "n": s["name"], "sid": sid, "wh": wh or "", "city": s["city"],
        "expo": s["expo"], "visit": s["visit"], "buyers": s["buyers"], "newb": s["newb"],
        "orders": o, "pay": round(s["pay"]), "sub": round(s["sub"]),
        "c_cancel": s["c_cancel"], "late": s["late"], "oos": s["oos"], "quality": s["quality"],
        "bad": s["bad"], "comp": s["comp"], "im": s["im"], "im_1m": s["im_1m"],
        "score": round(s["s_score"] / o, 2) if o else 0, "dur": round(s["dur"] / s["days"], 1) if s["days"] else 0,
    })

daily_out = [[DAYS[d], v[0], v[1], v[2], v[3], round(v[4])] for d, v in sorted(daily.items())]


def pack19(dv):
    return [int(dv[0]), int(dv[1]), int(dv[2]), int(dv[3]), int(dv[4]),
            round(dv[5]), round(dv[6]), round(dv[7]), int(dv[8]), int(dv[9]),
            int(dv[10]), int(dv[11]), int(dv[12]), int(dv[13]), int(dv[14]),
            round(dv[15], 1), round(dv[16], 1), round(dv[17], 1), round(dv[18], 1)]


# 稀疏输出：{仓: {天序号: [19字段]}}，仅保留有流量的天
whd_out = {}
for wh, daymap in wh_daily.items():
    whd_out[wh] = {str(di): pack19(dv) for di, dv in sorted(daymap.items())}

std_out = {}
for sid, daymap in sid_daily.items():
    o = {}
    for di, dv in sorted(daymap.items()):
        o[str(di)] = [int(dv[0]), int(dv[1]), int(dv[2]), int(dv[3]), int(dv[4]),
                      round(dv[5]), int(dv[10]), int(dv[8]), round(dv[17], 1)]
    std_out[sid] = o

ordd_out = {wh: {str(di): x for di, x in sorted(m.items())} for wh, m in ordd.items()}

total = {"expo": sum(w["expo"] for w in out_wh), "visit": sum(w["visit"] for w in out_wh),
         "buyers": sum(w["buyers"] for w in out_wh), "newb": sum(w["newb"] for w in out_wh),
         "porders": sum(w["porders"] for w in out_wh), "rev": sum(w["rev"] for w in out_wh),
         "ord": sum(w["ord"] for w in out_wh), "pay": sum(w["pay"] for w in out_wh),
         "sub": sum(w["sub"] for w in out_wh), "stores": sum(w["stores"] for w in out_wh)}
payload = {
    "meta": {
        "period": DAYS[0] + " ~ " + DAYS[-1],
        "platform": "美团闪购",
        "wh": len(out_wh), "stores": len(stores), "store_matched": len(stores) - unmatch["stores"],
        "unmatch": {"stores": unmatch["stores"], "expo": unmatch["expo"], "orders": unmatch["porders"], "names": unmatch["names"]},
        "total": total,
        "note": "营收/单量取订单底表（美团闪购、归属表内仓、剔除黑名单店铺）；曝光/入店/下单/新客/转化率取平台门店明细，按商家ID关联渠道店铺ID",
    },
    "days": DAYS,
    "wh": sorted(out_wh, key=lambda x: (-x.get("expo", 0), -x["rev"])),
    "store": sorted(out_store, key=lambda x: -x["expo"]),
    "daily": daily_out,
    "whd": whd_out,          # 仓级按日（稀疏 {天序号: [19字段]}）
    "ordd": ordd_out,        # 底表美团按日（稀疏 {天序号: [单量, 营收元]}）
    "std": std_out,          # 门店级按日（稀疏 {天序号: [9字段]}）
}
os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
json.dump(payload, open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
print("已写出", OUT_JSON)

# ── 体检输出 ──
def fmt_i(v):
    return format(int(round(v)), ",")


print("仓 %d | 门店 %d（匹配 %d / 未匹配 %d）" % (len(out_wh), len(stores), payload["meta"]["store_matched"], unmatch["stores"]))
print("合计：曝光 %s | 入店 %s | 下单 %s | 新客 %s | 平台单 %s | 底表单 %s | 营收 ¥%s" % (
    fmt_i(total["expo"]), fmt_i(total["visit"]), fmt_i(total["buyers"]), fmt_i(total["newb"]),
    fmt_i(total["porders"]), fmt_i(total["ord"]), fmt_i(total["rev"])))
print("入店转化 %.2f%% | 下单转化 %.2f%% | 新客占比 %.1f%% | 折扣率 %.1f%% | 补贴强度 %.1f%%" % (
    total["visit"] / total["expo"] * 100, total["buyers"] / total["visit"] * 100,
    total["newb"] / total["buyers"] * 100,
    total["pay"] / sum(w["price"] for w in out_wh) * 100 if sum(w["price"] for w in out_wh) else 0,
    total["sub"] / sum(w["price"] for w in out_wh) * 100 if sum(w["price"] for w in out_wh) else 0))
print("未匹配商家 曝光 %s / 单 %s（%.1f%% 曝光）" % (
    fmt_i(unmatch["expo"]), fmt_i(unmatch["porders"]), unmatch["expo"] / (total["expo"] + unmatch["expo"]) * 100))
# 按月对账：底表 vs 平台
bym_b = collections.Counter(); bym_p = collections.Counter()
for wh, m in ordd_out.items():
    for di, x in m.items():
        bym_b[DAYS[int(di)][:7]] += x[0]
for wh, m in whd_out.items():
    for di, dv in m.items():
        bym_p[DAYS[int(di)][:7]] += dv[4]
for mo in sorted(set(bym_b) | set(bym_p)):
    b, p = bym_b.get(mo, 0), bym_p.get(mo, 0)
    flag = "OK" if abs(b - p) <= max(20, p * 0.02) else "≠≠"
    print("  %s 月 底表 %7d vs 平台 %7d  %s" % (mo, b, p, flag))
print("样本 TOP5 营收仓:")
for w in payload["wh"][:5]:
    print("  %-26s 曝光%8s 入店%7s 单%6s 营收¥%10s 入店转化%5.1f%% 下单转化%5.1f%%" % (
        w["name"][:26], fmt_i(w["expo"]), fmt_i(w["visit"]), fmt_i(w["ord"]), fmt_i(w["rev"]),
        w["visit"] / w["expo"] * 100 if w["expo"] else 0,
        w["buyers"] / w["visit"] * 100 if w["visit"] else 0))
