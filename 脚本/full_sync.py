#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""跨平台「完整同步」脚本（Windows / macOS 通用，仅用标准库）
================================================================
把过去需要人工分步执行的链路代码化，任何一台机器都能独立完成
「导入本地 CSV → 重建全部数据 → 推送上线」：

  1. 对齐云端历史（pull --rebase，带 PAT 认证；失败则 fetch + reset --soft）
  2. update_dashboard.py        存档 + 模板内嵌
  3. 抽取模板 __EMBEDDED_GZ__   → dashboard-data.bin；模板复制为 index.html
  4. build_traffic.py / build_traffic_page.py / build_goods.py
  5. git add + commit → pin_cdn_version()（钉 CDN 不可变地址）
  6. 推送：git push 优先，失败自动回退 api_push.py（REST API，绕过被拦的 443）

用法：
  python3 脚本/full_sync.py                 # 完整流程（推荐）
  python3 脚本/full_sync.py --no-pull      # 不拉云端（仅离线本地重建）
  python3 脚本/full_sync.py --no-push      # 只生成，不提交/不推送
  python3 脚本/full_sync.py --push=api     # 强制走 REST API 推送
  RETAIL_DATA_DIR=~/我的底表 python3 脚本/full_sync.py   # 指定 CSV 目录

注意：本脚本不会调用 sync.gen_external_data()，因为它会清空 index.html 的内嵌数据。
"""
import base64
import datetime
import gzip
import io
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT_DIR = os.path.join(HERE, "脚本")
sys.path.insert(0, SCRIPT_DIR)

PY = sys.executable
CFG_REL = os.path.join("配置与映射", "sync_config.json")

# 每次更新需要提交的文件（数据产物 + 生成源 + 存档）
COMMIT_FILES = [
    os.path.join("历史数据", "历史数据存档.json"),
    os.path.join("看板页面", "retail-dashboard.html"),
    "index.html",
    "dashboard-data.bin",
    "traffic.html",
    "traffic-data.bin",
    "goods-data.bin",
]
# pull 前需要丢弃本地改动的数据产物（随后由流水线重新生成）
DATA_PRODUCTS = [
    os.path.join("历史数据", "历史数据存档.json"),
    "dashboard-data.bin",
    "traffic-data.bin",
    "goods-data.bin",
]


def log(msg):
    print("[同步] " + msg, flush=True)


def _config():
    """复用 sync.py 的配置读取：环境变量 RETAIL_SYNC_PAT 优先，其次本地 json。"""
    try:
        import sync
        cfg = sync.load_config()
        token = sync.get_pat(cfg)
    except Exception as e:
        log("读取同步配置失败（按无令牌模式继续）：%s" % e)
        cfg, token = {"remote": "origin", "branch": "main", "remote_url": ""}, ""
    return cfg, token


def git(*args, check=True):
    cmd = ["git", "-C", HERE] + list(args)
    p = subprocess.run(cmd, cwd=HERE, stdout=subprocess.PIPE,
                       stderr=subprocess.STDOUT)
    out = p.stdout.decode("utf-8", "replace")
    if check and p.returncode != 0:
        raise RuntimeError("git %s 失败：%s" % (" ".join(args), out[:400]))
    return p.returncode == 0, out


def run_py(rel, *args):
    p = subprocess.run([PY, os.path.join(SCRIPT_DIR, rel)] + list(args), cwd=HERE)
    if p.returncode != 0:
        raise RuntimeError("%s 执行失败（返回码 %d）" % (rel, p.returncode))


def auth_url(cfg, token):
    """把令牌注入远程 URL，避免交互式认证 / credential helper 依赖。"""
    url = cfg.get("remote_url") or ""
    if not token or not url.startswith("https://github.com/"):
        return ""
    return url.replace("https://", "https://x-access-token:%s@" % token, 1)


def pull_remote(cfg, token):
    """对齐云端历史。数据产物先丢弃（后面会从本机 CSV 重新生成）。"""
    log("对齐云端历史…")
    for f in DATA_PRODUCTS:
        git("checkout", "--", f, check=False)
    # 页面源码若有人工改动，先自动提交保护（与 sync.py 行为一致）；
    # 无改动时绝不能 git add，否则 rebase 会因暂存区非空而失败。
    for f in ["看板页面/retail-dashboard.html", "index.html", "traffic.html", "hub.html"]:
        changed, _ = git("diff", "--quiet", "--", f, check=False)
        cached, _ = git("diff", "--quiet", "--cached", "--", f, check=False)
        if not changed or not cached:
            git("add", "--", f, check=False)
            git("commit", "-m", "页面更新：full_sync 自动保护未提交的页面改动", check=False)
    ok, out = git("-c", "http.lowSpeedLimit=1000", "-c", "http.lowSpeedTime=30",
                  "pull", "--rebase", cfg.get("remote", "origin"), cfg.get("branch", "main"),
                  check=False)
    if ok:
        log("已拉取云端最新。")
        return
    log("pull --rebase 未成功，尝试 fetch + 软重置：%s" % out.strip().splitlines()[-1:])
    remote_ref = "%s/%s" % (cfg.get("remote", "origin"), cfg.get("branch", "main"))
    url = auth_url(cfg, token)
    if url:
        git("-c", "http.lowSpeedLimit=1000", "fetch", url,
            "%s:refs/remotes/%s" % (cfg.get("branch", "main"), remote_ref), check=False)
    else:
        git("fetch", cfg.get("remote", "origin"), cfg.get("branch", "main"), check=False)
    ok2, _ = git("rev-parse", "--verify", remote_ref, check=False)
    if ok2:
        git("reset", "--soft", remote_ref)
        for f in DATA_PRODUCTS:
            git("checkout", "--", f, check=False)
        log("已对齐云端历史（保留工作区改动）。")
    else:
        log("未取得远端引用，按离线模式继续（请勿同时双机更新）。")


def rebuild_all():
    log("1/5 处理订单底表并合并历史存档…")
    run_py("update_dashboard.py", "--no-open")

    log("2/5 抽取数据包并同步入口页…")
    tpl = io.open(os.path.join(HERE, "看板页面", "retail-dashboard.html"),
                  encoding="utf-8", errors="ignore").read()
    m = re.search(r'var __EMBEDDED_GZ__ = "([^"]+)"', tpl)
    if not m:
        raise RuntimeError("模板中未找到 __EMBEDDED_GZ__，无法生成 dashboard-data.bin")
    raw = base64.b64decode(m.group(1))
    with open(os.path.join(HERE, "dashboard-data.bin"), "wb") as f:
        f.write(raw)
    shutil.copyfile(os.path.join(HERE, "看板页面", "retail-dashboard.html"),
                    os.path.join(HERE, "index.html"))
    d = json.loads(gzip.decompress(raw).decode("utf-8"))
    days = d.get("DATA_DAYS", [])
    log("    数据截至 %s｜%d 天｜%d 仓有数据 / %d 仓归属表"
        % (days[-1] if days else "?", len(days),
           len(d.get("WAREHOUSES", [])), len(d.get("WAREHOUSES_ALL", []))))

    log("3/5 重建流量数据与页面…")
    run_py("build_traffic.py")
    run_py("build_traffic_page.py")

    log("4/5 重建商品数据…")
    run_py("build_goods.py")
    return days[-1] if days else "?"


def commit_and_pin(last_day):
    for f in COMMIT_FILES:
        git("add", "--", f)
    ok, out = git("diff", "--cached", "--quiet", check=False)
    if ok:
        log("数据与昨日一致，跳过提交（仍会检查钉版）。")
    else:
        msg = "数据更新至 %s（full_sync.py 跨平台链路）" % last_day
        git("commit", "-m", msg)
        log("已提交：%s" % msg)
    # 复用 sync.py 的钉版逻辑：把 5 个页面的 CDN 地址钉到数据提交 sha
    import sync
    sync.pin_cdn_version()


def push_all(cfg, token, mode):
    url = auth_url(cfg, token)
    target = url or cfg.get("remote", "origin")
    branch = cfg.get("branch", "main")
    if mode in ("auto", "git"):
        ok, out = git("push", "--set-upstream", target, "HEAD:%s" % branch, check=False)
        if ok:
            log("git push 成功。")
        elif mode == "git":
            raise RuntimeError("git push 失败：%s" % out[-400:])
        else:
            log("git push 失败（可能是 443 被拦），改用 REST API 推送…")
    if mode == "api" or (mode == "auto" and not ok):
        env = os.environ.copy()
        # Windows 本机代理偶发 502：REST 推送默认绕开代理；Mac 如需代理可设 RETAIL_KEEP_PROXY=1
        if not env.get("RETAIL_KEEP_PROXY"):
            for k in ("http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY"):
                env.pop(k, None)
        p = subprocess.run([PY, os.path.join(SCRIPT_DIR, "api_push.py")],
                           cwd=HERE, env=env)
        if p.returncode != 0:
            raise RuntimeError("REST API 推送也失败，请检查令牌与网络。")
    # 让本地远端跟踪引用与真实远端一致，避免下次 pull/reset 回退
    git("update-ref", "refs/remotes/%s/%s" % (cfg.get("remote", "origin"), branch),
        "HEAD", check=False)


def main():
    args = sys.argv[1:]
    no_pull = "--no-pull" in args
    no_push = "--no-push" in args
    mode = "auto"
    for a in args:
        if a.startswith("--push="):
            mode = a.split("=", 1)[1]
    if mode not in ("auto", "git", "api"):
        sys.exit("--push 只支持 auto / git / api")

    ok, _ = git("rev-parse", "--is-inside-work-tree", check=False)
    if not ok:
        sys.exit("当前目录不是 git 仓库，请先 git clone 本项目。")
    cfg, token = _config()
    if not token:
        log("未配置 RETAIL_SYNC_PAT / sync_config.json；将依赖系统 git 凭据。")

    log("=== 完整同步开始（%s）===" % ("macOS/Linux" if os.name != "nt" else "Windows"))
    if not no_pull:
        pull_remote(cfg, token)
    last_day = rebuild_all()
    if no_push:
        log("已按要求跳过提交与推送，本地产物已生成。")
        return
    commit_and_pin(last_day)
    push_all(cfg, token, mode)
    log("=== 完成 ===  线上：https://l-rom0609.github.io/retail-monitor/")


if __name__ == "__main__":
    main()
