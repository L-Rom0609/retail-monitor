# -*- coding: utf-8 -*-
"""美团流量页的公共样式与工具函数（被 build_traffic_page.py 引用）"""

COMMON_CSS = r"""
:root{
  --bg:#F4F7FC;--card:#FFFFFF;--ink:#0B2447;--ink2:#33475F;--muted:#6B7A90;--muted2:#93A1B5;
  --line:#E4EBF5;--brand:#2563EB;--brand-soft:#EBF3FE;--mt:#FFC300;--mt-ink:#7A5A00;
  --up:#C0281C;--up-soft:#FDECEA;--down:#0E7A46;--down-soft:#E9F7F0;
  --warn:#DE8A0B;--warn-soft:#FFF6E6;--hi:#C0281C;--hi-soft:#FDECEA;
  --r:14px;--sh:0 2px 10px rgba(11,36,71,.06);
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{background:var(--bg);color:var(--ink);
  font:13px/1.5 -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei","Noto Sans SC",sans-serif;
  -webkit-font-smoothing:antialiased;}
.tabnum{font-variant-numeric:tabular-nums;font-feature-settings:"tnum";}
.wrap{max-width:1480px;margin:0 auto;padding:18px 20px 40px;}
/* 顶栏 */
.top{display:flex;align-items:flex-end;gap:14px;flex-wrap:wrap;margin-bottom:14px;}
.top h1{font-size:21px;line-height:1.25;margin:0;letter-spacing:-.2px;}
.top h1 em{font-style:normal;color:var(--brand);}
.top .sub{font-size:12px;color:var(--muted);margin-top:4px;}
.top .right{margin-left:auto;display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
.pill{display:inline-flex;align-items:center;gap:6px;height:26px;padding:0 11px;border-radius:999px;
  font-size:11.5px;font-weight:600;background:#fff;border:1px solid var(--line);color:var(--ink2);}
.pill.preview{background:#FFF7DF;border-color:#F3DEA0;color:#7A5A00;}
.pill.mt{background:#FFFBE8;border-color:#F5E3A6;color:#7A5A00;}
.dot{width:7px;height:7px;border-radius:50%;background:currentColor;display:inline-block;}
/* 卡片 */
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--sh);}
.card-h{display:flex;align-items:center;gap:10px;padding:13px 16px 0;}
.card-h h2{font-size:13.5px;margin:0;font-weight:700;letter-spacing:.1px;}
.card-h .hint{font-size:11.5px;color:var(--muted2);margin-left:auto;}
/* KPI */
.kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin-bottom:12px;}
.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:11px 13px 12px;box-shadow:var(--sh);position:relative;overflow:hidden;}
.kpi::after{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--kc,var(--brand));opacity:.85;}
.kpi .k{font-size:11.5px;color:var(--muted);display:flex;align-items:center;gap:5px;}
.kpi .v{font-size:22px;font-weight:750;letter-spacing:-.4px;margin-top:4px;line-height:1.1;}
.kpi .s{font-size:11px;color:var(--muted2);margin-top:3px;}
.kpi .s b{color:var(--ink2);font-weight:700;}
/* 表格 */
.tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:10px 14px;border-bottom:1px solid var(--line);}
.tools select,.tools input{height:30px;border:1px solid var(--line);border-radius:8px;background:#fff;
  font:12.5px/1 inherit;color:var(--ink);padding:0 9px;outline:none;}
.tools input{width:190px;}
.tools select:focus,.tools input:focus{border-color:#9DC0F8;box-shadow:0 0 0 3px rgba(37,99,235,.10);}
.seg{display:inline-flex;background:#F1F5FB;border-radius:8px;padding:2px;gap:2px;}
.seg button{border:0;background:transparent;font:12px/1 inherit;color:var(--muted);padding:6px 11px;border-radius:6px;cursor:pointer;}
.seg button.on{background:#fff;color:var(--ink);font-weight:700;box-shadow:0 1px 3px rgba(11,36,71,.12);}
.count{margin-left:auto;font-size:11.5px;color:var(--muted2);}
table{width:100%;border-collapse:separate;border-spacing:0;}
thead th{position:sticky;top:0;z-index:2;background:#F7FAFF;font-size:11px;color:var(--muted);
  font-weight:700;text-align:right;padding:8px 10px;white-space:nowrap;border-bottom:1px solid var(--line);cursor:pointer;user-select:none;}
thead th.l{text-align:left;}
thead th:hover{color:var(--brand);}
thead th .ar{opacity:.35;margin-left:3px;font-size:9px;}
thead th.on .ar{opacity:1;color:var(--brand);}
tbody td{padding:8px 10px;font-size:12.5px;text-align:right;border-bottom:1px solid #F1F5FB;white-space:nowrap;}
tbody td.l{text-align:left;}
tbody tr:hover td{background:#F8FBFF;}
.wh{font-weight:600;color:var(--ink);}
.mgr{display:inline-block;font-size:10.5px;color:var(--muted);background:#F3F6FB;border-radius:5px;padding:1px 6px;margin-left:6px;}
.rate{font-weight:700;}
.rate.low{color:var(--up);}
.rate.good{color:var(--down);}
.bar-cell{display:flex;align-items:center;gap:7px;justify-content:flex-end;}
.bar{width:54px;height:5px;border-radius:3px;background:#EEF3FA;overflow:hidden;flex:none;}
.bar i{display:block;height:100%;border-radius:3px;background:linear-gradient(90deg,#7FB0F5,#2563EB);}
.badge{display:inline-flex;align-items:center;gap:3px;height:19px;padding:0 7px;border-radius:6px;
  font-size:10.5px;font-weight:700;white-space:nowrap;}
.badge.hi{background:var(--hi-soft);color:var(--hi);}
.badge.mid{background:var(--warn-soft);color:#95620A;}
.badge.ok{background:var(--down-soft);color:var(--down);}
.badge.none{background:#F3F6FB;color:var(--muted2);font-weight:600;}
.risk-cell{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end;}
.empty{padding:40px 0;text-align:center;color:var(--muted2);font-size:12.5px;}
.note{font-size:11.5px;color:var(--muted2);line-height:1.7;margin-top:10px;}
.note b{color:var(--ink2);}
@media (max-width:1100px){.kpis{grid-template-columns:repeat(3,minmax(0,1fr));}}
@media (max-width:700px){
  .wrap{padding:12px 10px 30px;}
  .top h1{font-size:18px;}
  .kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;}
  .kpi .v{font-size:19px;}
  .tools input{width:100%;}
  .tools .count{width:100%;text-align:left;}
  .tbl-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;}
  table{min-width:920px;}
}
"""

COMMON_JS = r"""
const fmtI = v => Math.round(v||0).toLocaleString('zh-CN');
const fmtM = v => { v = Math.round(v||0); return v >= 100000000 ? '¥'+(v/100000000).toFixed(2)+'亿' : (v >= 10000 ? '¥'+(v/10000).toFixed(1)+'万' : '¥'+fmtI(v)); };
const fmtAov = v => '¥'+(v||0).toFixed(1);
const pc = (a,b) => (b ? a/b*100 : 0);
const p1 = v => (v||0).toFixed(1)+'%';
const esc = s => String(s==null?'':s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const median = arr => { const a = arr.filter(v=>isFinite(v)).sort((x,y)=>x-y); return a.length ? a[Math.floor(a.length/2)] : 0; };

function M(w){
  const o = w.porders || 0;
  return {
    ir: pc(w.visit, w.expo), cr: pc(w.buyers, w.visit), nr: pc(w.newb, w.buyers),
    aov: w.ord ? w.rev/w.ord : 0,
    oos: pc(w.oos,o), cc: pc(w.c_cancel,o), late: pc(w.late,o),
    bad: pc(w.bad,o), comp: pc(w.comp,o),
    im1: w.im ? pc(w.im_1m, w.im) : null,
    disc: pc(w.pay, w.price), subr: pc(w.sub, w.price),
    online: w.avg_plan ? pc(w.avg_online, w.avg_plan) : null
  };
}

/* 运营风险引擎：只保留真正影响店铺权重/转化的字段 */
function risks(w, base){
  const m = M(w), r = [];
  if (!w.porders) return r;
  if (m.oos >= 3) r.push(['缺货退款 '+p1(m.oos), 'hi', '缺货退款率偏高：平台判责直接扣店铺分、掉搜索权重，还会拉低下单转化']);
  else if (m.oos >= 1) r.push(['缺货退款 '+p1(m.oos), 'mid', '出现缺货退款：建议按动销补货，高动销品设置安全库存']);
  if (m.cc >= 3) r.push(['商家取消 '+p1(m.cc), 'hi', '商家原因取消多：直接损伤履约分与门店排名']);
  else if (m.cc >= 1) r.push(['商家取消 '+p1(m.cc), 'mid', '偶发商家取消：排查接单人力与库存同步']);
  if (m.late >= 0.5) r.push(['超时未接 '+p1(m.late), 'hi', '超时未接单：高峰期接单能力不足，直接影响曝光']);
  if (m.bad >= 2) r.push(['差评 '+p1(m.bad), 'hi', '差评率偏高：口碑分下滑，进店与下单双杀']);
  else if (m.bad > 0) r.push(['差评 '+p1(m.bad), 'mid', '有差评：建议当天回访处理']);
  if (m.comp >= 0.8) r.push(['投诉 '+p1(m.comp), 'hi', '投诉率高：可能触发平台介入与扣分']);
  if (m.im1 !== null && m.im1 < 70) r.push(['IM 响应 '+p1(m.im1), 'hi', '1 分钟响应率低于 70%：服务分受损，咨询客户流失']);
  else if (m.im1 !== null && m.im1 < 90) r.push(['IM 响应 '+p1(m.im1), 'mid', '1 分钟响应率偏低：建议开启自动回复']);
  if (w.s_score && w.s_score < 4.0) r.push(['评分 '+w.s_score.toFixed(1), 'hi', '商家评分低于 4.0：搜索排序与进店意愿同时受压']);
  else if (w.s_score && w.s_score < 4.6) r.push(['评分 '+w.s_score.toFixed(1), 'mid', '商家评分偏低：关注差评来源']);
  if (m.online !== null && m.online < 80) r.push(['在线时长 '+p1(m.online), 'mid', '在线时长不足预设：每天少营业即少曝光']);
  if (m.ir < base.ir * 0.7) r.push(['进店转化偏低 '+p1(m.ir), 'mid', '曝光有了但不进店：头图/店名/价格带/评分是主因']);
  if (m.cr < base.cr * 0.7) r.push(['下单转化偏低 '+p1(m.cr), 'mid', '进店了不下单：详情页、满减门槛、库存与配送时长需排查']);
  return r;
}
const riskScore = list => list.reduce((s,x)=>s+(x[1]==='hi'?3:1),0);
function riskBadges(list, max){
  if (!list.length) return '<span class="badge none">正常</span>';
  const hi = list.filter(x=>x[1]==='hi').length;
  let out = '';
  const show = list.slice(0, max||3);
  show.forEach(x => { out += '<span class="badge '+x[1]+'" title="'+esc(x[2])+'">'+esc(x[0])+'</span>'; });
  if (list.length > show.length) out += '<span class="badge none">+'+(list.length-show.length)+'</span>';
  return out;
}
function quantile(arr, q){ const a = arr.filter(v=>isFinite(v)).sort((x,y)=>x-y); return a.length ? a[Math.min(a.length-1, Math.floor(a.length*q))] : 0; }
"""
