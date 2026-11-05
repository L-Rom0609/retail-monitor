# -*- coding: utf-8 -*-
"""
双端 Git 同步包装器（Windows / macOS 通用，仅用标准库）
===================================================
把本仓库当作「即时零售看板」的唯一真源：

  1. 拉取云端最新归档 + 归属表（git pull --rebase）
  2. 运行 update_dashboard.py 处理本地 CSV，把结果合并进历史数据仓库
  3. 提交并把最新归档 + 归属表推回云端（git commit + push）
  4. 若推送被另一台机器抢先（冲突），自动 rebase 后重跑合并再推，最多重试几次

这样 Mac 与 Windows 都能独立更新看板，且两端数据合并不互相覆盖。
凭据（PAT）只从环境变量 RETAIL_SYNC_PAT 读取，绝不写进仓库。

用法：
  python sync.py [--no-open] [其它传给 update_dashboard.py 的参数]
  （--no-open 等参数会原样透传给 update_dashboard.py）
"""
import io
import os
import sys
import json
import hashlib
import re
import shutil
import subprocess
from base64 import b64decode as _b64decode
import urllib.request as _urlreq

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根目录（本脚本位于 脚本/ 子文件夹内）

# 需要纳入版本控制、随仓库在双端保持一致的「真源」文件
TRACKED = [
    "脚本/update_dashboard.py",
    "脚本/sync.py",
    "历史数据/历史数据存档.json",
    "配置与映射/仓名-负责人归属关系表.xlsx",
    "配置与映射/商品成本表.xlsx",
    "看板页面/retail-dashboard.html",
    "index.html",
    "配置与映射/月度目标配置.json",
    "dashboard-data.bin",
    "goods.html",
    "goods-data.bin",
    "traffic.html",
    "traffic-data.bin",
    "脚本/build_traffic.py",
    "脚本/build_traffic_page.py",
    "脚本/traffic_common.py",
    "脚本/build_goods.py",
    "配置与映射/赠品与提示行名单.txt",
    "README.md",
    ".gitignore",
    "配置与映射/sync_config.example.json",
    "一键更新看板.bat",
    "脚本/update.sh",
]

# 只有「数据产物」在 pull 前丢弃本地改动（由流水线重新生成）；
# 脚本/配置类文件保留本地改动，避免把你改的代码冲掉。
# 注意：retail-dashboard.html / index.html 既是页面源码又是生成产物，
# 绝不能在这里 checkout 丢弃——否则页面改动（如本次的汇总卡升级）会被 pull 冲掉。
# 它们由 git_pull() 里的「页面改动自动提交保护」单独处理。
DATA_FILES = [
    "历史数据/历史数据存档.json",
    "dashboard-data.bin",
]

# 页面源码型文件：本地有未提交改动 = 页面被人工修改过，pull 前先自动提交保护。
PAGE_FILES = [
    "看板页面/retail-dashboard.html",
    "index.html",
    "traffic.html",
]

# 本地配置（可选，git 忽略）：sync_config.json
CONFIG_PATH = os.path.join(HERE, "配置与映射", "sync_config.json")
MAX_PUSH_RETRY = 5

# ── 历史自动压缩（控制仓库体积）────────────────────────────────────────────
# 每次更新都会新增一版「历史数据存档.json（16MB）+ 页面 + 数据文件」，
# Git 压缩后每次约增加 2~3MB。攒到一定提交数就自动把历史压成一条提交，
# 体积回到几 MB。可用环境变量临时覆盖，便于测试：
#   RETAIL_SQUASH_MIN_COMMITS / RETAIL_SQUASH_MIN_MB
def _env_int(name, default):
    try:
        return int(os.environ.get(name, "") or default)
    except ValueError:
        return default


SQUASH_MIN_COMMITS = _env_int("RETAIL_SQUASH_MIN_COMMITS", 60)   # ≈ 两个月（每天一次更新）
SQUASH_MIN_MB = _env_int("RETAIL_SQUASH_MIN_MB", 200)            # 或 .git 超过 200MB


def log(msg):
    print("[同步] " + msg)


def load_config():
    cfg = {
        "remote": "origin",
        "branch": "main",
        "remote_url": os.environ.get("RETAIL_SYNC_REMOTE", ""),
        "pat_env": "RETAIL_SYNC_PAT",
        # 也可把令牌直接写进 sync_config.json 的 pat 字段（该文件已被 .gitignore 忽略，
        # 不会推向云端）。这样双击 一键更新看板.bat 即可免手工设环境变量。
        "pat": "",
    }
    if os.path.isfile(CONFIG_PATH):
        try:
            with io.open(CONFIG_PATH, encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict):
                cfg.update({k: d[k] for k in cfg if k in d})
        except Exception as e:
            log("读取 sync_config.json 失败（用默认配置）：%s" % e)
    return cfg


def _find_git():
    g = shutil.which("git") or shutil.which("git.exe")
    if g:
        return g
    # WorkBuddy 自带的 PortableGit（实际布局：versions/<版本>/cmd/git.exe 与
    # versions/<版本>/mingw64/bin/git.exe，版本根目录下没有 bin/）
    pats = [
        os.path.join(os.path.expanduser("~"), ".workbuddy", "binaries", "PortableGit",
                     "versions", "*", "cmd", "git.exe"),
        os.path.join(os.path.expanduser("~"), ".workbuddy", "binaries", "PortableGit",
                     "versions", "*", "mingw64", "bin", "git.exe"),
        os.path.join(os.path.expanduser("~"), ".workbuddy", "binaries", "PortableGit",
                     "versions", "*", "bin", "git.exe"),
        os.path.join(os.path.expanduser("~"), ".workbuddy", "binaries", "PortableGit",
                     "versions", "*", "bin", "git"),
    ]
    import glob as _g
    for p in pats:
        for f in sorted(_g.glob(p), reverse=True):
            if os.path.isfile(f):
                return f
    return None


GIT = _find_git()


def git(*args, check=True, capture=True):
    if not GIT:
        raise RuntimeError("本机未找到 git，请先安装 Git 并加入 PATH。")
    cmd = [GIT, "-C", HERE] + list(args)
    env = dict(os.environ)
    env["GIT_TERMINAL_PROMPT"] = "0"   # 网络不通时禁止弹出凭据等待输入，避免窗口卡死
    return subprocess.run(cmd, stdout=subprocess.PIPE if capture else None,
                          stderr=subprocess.STDOUT if capture else None,
                          cwd=HERE, env=env)


def git_ok(*args):
    r = git(*args, check=False)
    return r.returncode == 0, (r.stdout or b"").decode("utf-8", "replace").strip()


def auth_url(base, pat):
    """把 PAT 注入 https 仓库地址（仅当还没有凭据时）。
    GitHub 的 fine-grained / classic 令牌都能用 x-access-token 作为用户名。"""
    if not pat:
        return base
    tail = base.split("://", 1)[-1]
    if "@" in tail:
        return base
    if tail.startswith("github.com"):
        return "https://x-access-token:%s@%s" % (pat, tail)
    return base.replace("://", "://" + pat + "@", 1)


def get_pat(cfg):
    """优先读环境变量，其次读 sync_config.json 的 pat 字段。"""
    return os.environ.get(cfg.get("pat_env", "RETAIL_SYNC_PAT"), "") or cfg.get("pat", "")


def has_remote(cfg):
    if not cfg.get("remote_url"):
        return False
    ok, _ = git_ok("remote", "get-url", cfg["remote"])
    return ok


def git_pull(cfg):
    """拉取云端最新；本地若有未提交改动先丢弃（由 update_dashboard.py 重新生成）。"""
    if not has_remote(cfg):
        log("未配置远程仓库，跳过拉取（将在本地生成/提交）。")
        return True
    # 自愈：上次同步若被中途打断（关窗口/断网/杀进程），仓库可能卡在 rebase 半途，
    # 不处理的话后续所有 git 操作都会失败。这里检测到残留 rebase 就中止还原。
    for rd in (os.path.join(HERE, ".git", "rebase-merge"),
               os.path.join(HERE, ".git", "rebase-apply")):
        if os.path.isdir(rd):
            log("检测到上次中断遗留的 rebase 状态，自动还原……")
            git("rebase", "--abort", check=False)
    # 丢弃对受跟踪数据产物的本地改动，保证工作区与远程一致、rebase 干净
    git("checkout", "--", *DATA_FILES, check=False)
    # 页面源码保护：retail-dashboard.html / index.html 有未提交改动 = 页面被改过，
    # 先自动提交，pull --rebase 后本地改动得以保留（曾被 checkout 冲掉过，属严重副作用）
    for f in PAGE_FILES:
        if not os.path.isfile(os.path.join(HERE, f)):
            continue
        chg, _ = git_ok("diff", "--quiet", "--", f)
        chg_cached, _ = git_ok("diff", "--quiet", "--cached", "--", f)
        if not chg or not chg_cached:
            git("add", "--", f)
            okc, _ = git_ok("commit", "-m", "页面更新：sync 自动保护未提交的页面改动")
            if okc:
                log("检测到未提交的页面改动，已自动提交保护：" + f)
    ok, out = git_ok("-c", "http.lowSpeedLimit=1000", "-c", "http.lowSpeedTime=30",
                     "pull", "--rebase", cfg["remote"], cfg["branch"])
    if ok:
        log("已拉取云端最新数据。")
        return True

    # 拉不动通常有两种原因：① 首次还没有上游；② 云端历史被压缩重写过（force push），
    # 本地提交与云端已无共同祖先。后者直接以云端为准重置即可——原始 CSV 仍在底表目录，
    # 重跑流水线会把本地数据重新合并回去，不会丢数。
    git_ok("-c", "http.lowSpeedLimit=1000", "-c", "http.lowSpeedTime=30",
           "fetch", cfg["remote"], cfg["branch"])
    target = "%s/%s" % (cfg["remote"], cfg["branch"])
    ok2, _ = git_ok("rev-parse", "--verify", target)
    if ok2:
        # 用 --soft：只把历史对齐到云端，工作区与暂存区原样保留。
        # （早期版本用 --hard 会把本机未提交的脚本改动一并冲掉，属严重副作用）
        ok3, _ = git_ok("reset", "--soft", target)
        if ok3:
            git("checkout", "--", *DATA_FILES, check=False)
            log("已对齐云端历史（云端被压缩过或与本地分叉）；本地未提交的改动已保留。")
            return True
    last = out.splitlines()[-1] if out else ""
    log("拉取完成（或尚无上游，可忽略）：%s" % last)
    return True


def run_update(passthrough):
    """运行 update_dashboard.py 处理本地 CSV 并合并进历史数据仓库。"""
    log("开始处理本地底表并更新看板数据…")
    r = subprocess.run([sys.executable, "脚本/update_dashboard.py", "--no-open"] + passthrough,
                       cwd=HERE)
    if r.returncode != 0:
        log("update_dashboard.py 执行失败（返回码 %d），已中止，不会推送。" % r.returncode)
        return False
    log("本地数据更新完成。")
    return True


def gen_external_data():
    """从生成的 HTML 中抽取最新数据，写出 dashboard-data.bin（gzip 二进制，列式编码），
    供 GitHub Pages / CDN 托管；看板启动会优先拉取它——这样“更新数据”只需 push
    该文件，无需重新发布应用（规避发布平台静态文件不刷新的问题）。
    同时把看板复制为 index.html（入口，去掉重复内嵌数据）。"""
    html_path = os.path.join(HERE, "看板页面", "retail-dashboard.html")
    out_bin = os.path.join(HERE, "dashboard-data.bin")
    old_txt = os.path.join(HERE, "dashboard-data.txt")
    idx_path = os.path.join(HERE, "index.html")
    try:
        h = io.open(html_path, encoding="utf-8", errors="ignore").read()
        m = re.search(r'var __EMBEDDED_GZ__ = "([^"]+)"', h)
        if not m:
            log("未在内嵌 HTML 中找到数据，跳过生成外部数据。")
            return
        b64 = m.group(1)
        # ★ 不再写 base64：base64 会让体积膨胀 33%，而浏览器 fetc 到的二进制
        #   可以直接交给 DecompressionStream 解压，省掉这一层编码。
        raw = _b64decode(b64)
        with open(out_bin, "wb") as f:
            f.write(raw)
        log("已生成外部数据文件 dashboard-data.bin（%.0f KB，gzip 二进制 / 列式，比原 base64 小 %.0f%%）"
            % (len(raw) / 1024.0, (1 - len(raw) / max(len(b64), 1)) * 100))
        # 旧版 base64 文本文件已废弃
        if os.path.exists(old_txt):
            try:
                os.remove(old_txt)
                log("已移除旧版 dashboard-data.txt（改用 .bin）")
            except Exception:
                pass
        # 复制为 index.html（入口）：清空内嵌数据，页面只拉外部文件，
        # 每次刷新少下载一份等量数据。本地 retail-dashboard.html 保留内嵌，断网可离线看。
        idx = re.sub(r'(var __EMBEDDED_GZ__ = ")[^"]*(")', r'\1\2', h, count=1)
        io.open(idx_path, "w", encoding="utf-8").write(idx)
        log("已生成 index.html（入口 %.0f KB，不含内嵌数据；数据由 dashboard-data.bin 提供）"
            % (len(idx.encode("utf-8")) / 1024.0))
    except Exception as e:
        log("生成外部数据失败：%s" % e)


def purge_cdn():
    """推送成功后清掉 jsDelivr 的 branch 缓存，保证同事刷新即见新数据。
    purge 失败不影响更新结果（最多滞后到 CDN 自然过期）。"""
    for name in ("dashboard-data.bin", "goods-data.bin", "traffic-data.bin"):
        url = "https://purge.jsdelivr.net/gh/L-Rom0609/retail-monitor@main/%s" % name
        try:
            req = _urlreq.Request(url, headers={"User-Agent": "retail-monitor-sync"})
            with _urlreq.urlopen(req, timeout=15) as r:
                body = r.read(300).decode("utf-8", "ignore")
            log("已请求刷新 CDN 缓存（%s）：%s" % (name, body.strip()[:90]))
        except Exception as e:
            log("刷新 CDN 缓存失败（%s，可忽略，稍后自动过期）：%s" % (name, e))


def pin_cdn_version():
    """把 index.html 里的 CDN 数据地址从 @main 钉到当前提交号。

    jsDelivr 对 @main 这类分支地址有约 12 小时缓存，purge 接口经常返回成功但
    节点仍是旧文件（表现为「数据已推上去，线上却还是旧数字」）。提交号地址是
    不可变的：每个新提交都是新 URL，CDN 必定回源，彻底绕开缓存问题。
    """
    ok, out = git_ok("rev-parse", "HEAD")
    if not ok:
        return
    sha = (out or "").strip().splitlines()[0][:12] if (out or "").strip() else ""
    if not sha:
        return
    import re as _re
    pat = _re.compile(r'@(?:main|[0-9a-f]{7,12})/dashboard-data\.bin')
    targets = [os.path.join(HERE, "index.html"),
               os.path.join(HERE, "看板页面", "retail-dashboard.html")]
    # 商品页同样走 jsDelivr：goods-data.bin 一并钉提交号（否则商品页吃 @main 12 小时缓存）
    pat_g = _re.compile(r'@(?:main|[0-9a-f]{7,12})/goods-data\.bin')
    changed = 0
    for p in targets:
        try:
            h = io.open(p, encoding="utf-8").read()
        except Exception:
            continue
        h2, n = pat.subn("@%s/dashboard-data.bin" % sha, h)
        if n:
            io.open(p, "w", encoding="utf-8").write(h2)
            git("add", os.path.relpath(p, HERE), check=False)
            changed += n
    try:
        g = os.path.join(HERE, "goods.html")
        h = io.open(g, encoding="utf-8").read()
        h2, n = pat_g.subn("@%s/goods-data.bin" % sha, h)
        if n:
            io.open(g, "w", encoding="utf-8").write(h2)
            git("add", "goods.html", check=False)
            changed += n
    except Exception:
        pass
    # 仓流量页同理：traffic-data.bin 也钉提交号
    pat_t = _re.compile(r'@(?:main|[0-9a-f]{7,12})/traffic-data\.bin')
    try:
        g = os.path.join(HERE, "traffic.html")
        h = io.open(g, encoding="utf-8").read()
        h2, n = pat_t.subn("@%s/traffic-data.bin" % sha, h)
        if n:
            io.open(g, "w", encoding="utf-8").write(h2)
            git("add", "traffic.html", check=False)
            changed += n
    except Exception:
        pass
    if not changed:
        return
    git("commit", "-m", "锁定 CDN 数据版本 @%s（不可变地址，绕开 @main 缓存）" % sha, check=False)
    log("已把数据地址钉到提交 %s（%d 处，含源文件本地预览；不再依赖 CDN 缓存刷新）" % (sha, changed))


def git_push(cfg):
    """提交受跟踪文件并推送到远程。返回 True 表示成功。"""
    git("add", *TRACKED, check=False)
    # 没有可提交改动就跳过
    ok, _ = git_ok("diff", "--cached", "--quiet")
    if ok:
        log("没有需要提交的变更。")
        return True
    msg = "数据更新 %s" % _now()
    git("commit", "-m", msg, check=False)
    pin_cdn_version()
    if not has_remote(cfg):
        log("未配置远程仓库，已提交到本地（待配置远程后推送）。")
        return True
    pat = get_pat(cfg)
    url = auth_url(cfg["remote_url"], pat)
    ok, out = git_ok("-c", "http.lowSpeedLimit=1000", "-c", "http.lowSpeedTime=30", "push", "--set-upstream", cfg["remote"], cfg["branch"])
    # 本机没配凭据助手时，直接用带令牌的地址推
    if not ok and pat:
        ok, out = git_ok("-c", "http.lowSpeedLimit=1000", "-c", "http.lowSpeedTime=30", "push", "--set-upstream", url, cfg["branch"])
        if not ok:
            ok, out = git_ok("-c", "http.lowSpeedLimit=1000", "-c", "http.lowSpeedTime=30", "push", url, cfg["branch"])
    if ok:
        log("已推送到云端，另一台机器拉取后即可获得最新数据。")
        purge_cdn()
        return True
    log("推送被拒绝（可能有其他人刚推送）：%s" % (out.splitlines()[-1] if out else ""))
    return False


def _commit_count():
    r = git("rev-list", "--count", "HEAD", check=False)
    try:
        return int((r.stdout or b"0").decode("utf-8", "replace").strip() or 0)
    except ValueError:
        return 0


def _git_dir_mb():
    total = 0
    base = os.path.join(HERE, ".git")
    for root, _dirs, files in os.walk(base):
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return total / 1048576.0


def squash_history(cfg):
    """把历史压成一条提交并强推云端（force push），随后清理本地悬空对象。
    只在「已成功推送、工作区干净」时调用，避免把没上传的数据压掉。"""
    url = auth_url(cfg["remote_url"], get_pat(cfg)) if cfg.get("remote_url") else ""
    if not url:
        log("未配置远程仓库，跳过历史压缩。")
        return False

    before_commits, before_mb = _commit_count(), _git_dir_mb()

    # 先把跟踪文件全部落盘提交，保证快照完整
    git("add", *TRACKED, check=False)
    ok, _ = git_ok("diff", "--cached", "--quiet")
    if not ok:
        git("commit", "-q", "-m", "压缩前快照 %s" % _now(), check=False)

    ok, out = git_ok("checkout", "--orphan", "__squash_tmp")
    if not ok:
        log("历史压缩失败（无法创建孤立分支）：%s" % (out.splitlines()[-1] if out else ""))
        return False
    git("add", "-A", check=False)
    git("commit", "-q", "-m", "历史压缩 %s（清理旧版本快照，减小仓库体积）" % _now(), check=False)
    # 用当前这条孤立提交覆盖 main（-M 会强制改名覆盖）
    ok, out = git_ok("branch", "-M", "main")
    if not ok:
        log("历史压缩失败（改名 main 失败）：%s" % (out.splitlines()[-1] if out else ""))
        return False

    ok, out = git_ok("push", "-f", "--set-upstream", cfg["remote"], cfg["branch"])
    if not ok:
        ok, out = git_ok("push", "-f", url, cfg["branch"])
    if not ok:
        log("历史压缩后强推失败（本地已完成压缩，下次运行会再试）：%s"
            % (out.splitlines()[-1] if out else ""))
        return False

    # 清掉旧版本留下的悬空对象，真正释放磁盘/体积
    git("reflog", "expire", "--expire=now", "--all", check=False)
    git("gc", "--prune=now", "--quiet", check=False)

    after_commits, after_mb = _commit_count(), _git_dir_mb()
    log("已压缩历史：提交 %d → %d 条，本地仓库 %.0fMB → %.0fMB（云端同体积）。"
        % (before_commits, after_commits, before_mb, after_mb))
    log("提示：另一台机器下次运行会自动对齐到压缩后的历史，无需手工处理。")
    return True


def maybe_squash(cfg):
    """超过阈值才压缩：提交数≥SQUASH_MIN_COMMITS 或 .git≥SQUASH_MIN_MB。"""
    n, mb = _commit_count(), _git_dir_mb()
    if n < SQUASH_MIN_COMMITS and mb < SQUASH_MIN_MB:
        log("仓库体积正常（%d 条提交 / 本地 %.0fMB，阈值 %d 条或 %dMB），跳过历史压缩。"
            % (n, mb, SQUASH_MIN_COMMITS, SQUASH_MIN_MB))
        return True
    log("仓库已达到压缩阈值（%d 条提交 / 本地 %.0fMB），开始自动压缩历史…" % (n, mb))
    return squash_history(cfg)


def _now():
    import time
    return time.strftime("%Y-%m-%d %H:%M")


def run_traffic_build():
    """重建仓流量页数据（traffic_payload.json → traffic.html + traffic-data.bin）。
    订单底表有新增日期后必须重跑，否则流量页看不到新数据。失败不阻断推送（流量页滞后一天可接受）。"""
    for cmd in (["build_traffic.py"], ["build_traffic_page.py"]):
        r = subprocess.run([sys.executable, os.path.join("脚本", cmd[0])],
                           cwd=HERE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        out = r.stdout.decode("utf-8", "replace")
        if r.returncode != 0:
            log("⚠ %s 失败（流量页数据未更新，不影响其他页面）：%s" % (cmd[0], out[-300:]))
            return False
    log("仓流量页数据已重建（traffic.html + traffic-data.bin）。")
    return True


def run_goods_build():
    """重建商品页数据（底表 → goods-data.bin）。
    商品页只吃 goods-data.bin，不重建就看不到新一天的订单/商品明细。
    失败不阻断推送（商品页滞后一天可接受，其余页面照常更新）。"""
    r = subprocess.run([sys.executable, os.path.join("脚本", "build_goods.py")],
                       cwd=HERE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    out = r.stdout.decode("utf-8", "replace")
    if r.returncode != 0:
        log("⚠ build_goods.py 失败（商品页数据未更新，不影响其他页面）：%s" % out[-300:])
        return False
    tail = [x for x in out.strip().splitlines() if x.startswith("订单 ")]
    log("商品页数据已重建（goods-data.bin）%s" % ("｜" + tail[-1] if tail else ""))
    return True


def main():
    passthrough = sys.argv[1:]
    # 历史压缩开关（不传给 update_dashboard.py）
    force_squash = "--squash" in passthrough
    no_squash = "--no-squash" in passthrough
    passthrough = [a for a in passthrough if a not in ("--squash", "--no-squash")]
    cfg = load_config()
    if not GIT:
        print("[同步] 错误：本机未找到 git，请安装 Git 后重试。")
        sys.exit(1)
    # 确保是 git 仓库
    ok, _ = git_ok("rev-parse", "--is-inside-work-tree")
    if not ok:
        log("当前目录还不是 git 仓库，请先执行：git init && git add -A && git commit -m init")
        sys.exit(1)

    log("=== 双端同步开始 ===")
    git_pull(cfg)

    if not run_update(passthrough):
        sys.exit(1)
    gen_external_data()
    run_traffic_build()
    run_goods_build()

    for attempt in range(1, MAX_PUSH_RETRY + 1):
        if git_push(cfg):
            break
        if attempt >= MAX_PUSH_RETRY:
            log("重试 %d 次仍推送失败，请稍后手动拉取最新后重试。" % MAX_PUSH_RETRY)
            sys.exit(1)
        log("第 %d 次重试：拉取最新并重新合并…" % attempt)
        git_pull(cfg)
        if not run_update(passthrough):
            sys.exit(1)
        gen_external_data()
        run_traffic_build()
        run_goods_build()

    # 推送成功后：按体积阈值自动压缩历史（--squash 强制、--no-squash 跳过）
    if force_squash:
        squash_history(cfg)
    elif not no_squash:
        maybe_squash(cfg)

    log("=== 双端同步完成 ===")
    log("已生成本地看板 index.html 与外部数据文件 dashboard-data.bin（同目录）。")
    log("线上更新机制：index.html 启动时先走 CDN（gcore / cdn.jsdelivr 并行竞速），失败回退同源。")
    log("因此『更新数据 = git push 本仓库』即可让线上自动刷新，无需重新发布；推送后会自动清 CDN 缓存。")
    log("首次部署需要：① 创建 GitHub 公开仓库并配置 RETAIL_SYNC_PAT；② 开启 GitHub Pages；")
    log("③ 我把仓库推上去。之后 Mac/Windows 任意一端运行 sync.py 即自动同步线上，你无需任何发布操作。")


if __name__ == "__main__":
    main()
