# -*- coding: utf-8 -*-
"""
api_push.py —— github.com:443 被网络拦截时的推送兜底。

github.com 直连经常 502/Empty reply，但 api.github.com 可用。
本脚本把「远端 HEAD → 本地 HEAD」之间的提交链，通过 GitHub REST API
(blobs → trees → commits → PATCH ref) 逐字节复刻推上去：

  - 内容取自本地 git 对象（git cat-file），message 保留原始字节（含尾部换行）；
  - author/committer 的姓名、邮箱、时间戳原样携带 → 生成的 commit sha 与本地完全一致；
  - 每个提交推完即比对返回 sha，不一致立即中止（避免远端分叉）；
  - 全部成功后远端 HEAD == 本地 HEAD，后续 git pull/push 恢复正常关系。

用法： python 脚本/api_push.py [提交说明（仅本地补提交时用）]
"""
import base64
import io
import json
import os
import subprocess
import sys
import urllib.request
import urllib.error

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(HERE, "脚本"))

API = "https://api.github.com/repos/L-Rom0609/retail-monitor"


def git(*args, binary=False):
    out = subprocess.run(["git"] + list(args), cwd=HERE, stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT)
    if out.returncode != 0:
        raise RuntimeError("git %s 失败: %s" % (" ".join(args), out.stdout[:300]))
    return out.stdout if binary else out.stdout.decode("utf-8", "replace")


def http(method, path, token, payload=None):
    url = API + path
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": "token " + token,
        "Accept": "application/vnd.github+json",
        "User-Agent": "retail-monitor-sync",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            body = r.read().decode("utf-8", "replace")
            return r.status, (json.loads(body) if body.strip() else {})
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")[:500]


def main():
    sys.path.insert(0, os.path.join(HERE, "脚本"))
    import importlib.util
    spec = importlib.util.spec_from_file_location("sync_mod", os.path.join(HERE, "脚本", "sync.py"))
    sync = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(sync)
    cfg = sync.load_config()
    token = sync.get_pat(cfg)
    if not token:
        sys.exit("未取得 PAT，无法走 API 推送")

    # 远端 HEAD
    code, d = http("GET", "/git/ref/heads/" + cfg["branch"], token)
    if code != 200:
        sys.exit("取远端 ref 失败 %s: %s" % (code, d))
    remote_sha = d["object"]["sha"]
    local_sha = git("rev-parse", "HEAD").strip()
    if remote_sha == local_sha:
        print("远端与本地一致（%s），无需推送。" % local_sha[:12])
        return

    # 待推链：remote..HEAD（旧→新）
    chain = git("log", "--reverse", "--format=%H", remote_sha + ".." + local_sha).split()
    if not chain:
        sys.exit("本地 HEAD 不包含远端 HEAD（历史分叉），请人工处理。")
    print("待推提交 %d 个：%s" % (len(chain), " → ".join(s[:7] for s in chain)))

    blob_cache = {}
    for csha in chain:
        raw = git("cat-file", "commit", csha, binary=True)
        head_b, msg_b = raw.split(b"\n\n", 1)
        headers = {}
        parents = []
        tree = None
        for line in head_b.decode("utf-8", "replace").splitlines():
            k, _, v = line.partition(" ")
            if k == "tree":
                tree = v
            elif k == "parent":
                parents.append(v)
            else:
                headers.setdefault(k, v)
        author_raw = headers["author"]
        committer_raw = headers["committer"]

        # 父提交的 tree
        parent_tree = git("rev-parse", "%s^{tree}" % parents[0]).strip() if parents else None

        # 变更文件列表
        dq = ["-c", "core.quotepath=off"]   # 中文路径不转义
        if parents:
            diff = git(*(dq + ["diff-tree", "-r", "--no-renames", "--no-commit-id", parents[0], csha]))
        else:
            diff = git(*(dq + ["diff-tree", "-r", "--no-renames", "--no-commit-id", "--root", csha]))
        entries = []
        for line in diff.splitlines():
            parts = line.split("\t", 1)
            if len(parts) != 2:
                continue
            meta, path = parts
            m1, m2, s1, s2, status = meta.split()
            if status == "D":
                entries.append({"path": path, "mode": "100644", "type": "blob", "sha": None})
                continue
            if s2 in blob_cache:
                bsha = blob_cache[s2]
            else:
                content = git("cat-file", "blob", s2, binary=True)
                code, d = http("POST", "/git/blobs", token,
                               {"content": base64.b64encode(content).decode(), "encoding": "base64"})
                if code != 201:
                    sys.exit("blob 上传失败 %s(%s): %s %s" % (path, s2[:8], code, d))
                bsha = d["sha"]
                if bsha != s2:
                    sys.exit("blob sha 不一致 %s: 本地 %s 远端 %s" % (path, s2, bsha))
                blob_cache[s2] = bsha
            entries.append({"path": path, "mode": m2, "type": "blob", "sha": bsha})

        code, d = http("POST", "/git/trees", token,
                       {"base_tree": parent_tree, "tree": entries} if entries
                       else {"base_tree": parent_tree, "tree": []})
        if code != 201:
            sys.exit("tree 创建失败 %s: %s %s" % (csha[:8], code, d))
        new_tree = d["sha"]

        def split_ident(s):
            # "Name <mail> 1793861827 +0800" → (name, email, ISO8601)
            lt = s.rindex(" <")
            gt = s.index(">", lt)
            name, email = s[:lt], s[lt + 2:gt]
            ts, tz = s[gt + 2:].split()
            sign = -1 if tz.startswith("-") else 1
            off = sign * (int(tz[1:3]) * 3600 + int(tz[3:5]) * 60)
            import datetime
            dt = datetime.datetime.utcfromtimestamp(int(ts) + off)
            iso = dt.strftime("%Y-%m-%dT%H:%M:%S") + tz[:3] + ":" + tz[3:]
            return name, email, iso

        an, ae, ai = split_ident(author_raw)
        cn, ce, ci = split_ident(committer_raw)
        code, d = http("POST", "/git/commits", token, {
            "message": msg_b.decode("utf-8"),
            "tree": new_tree,
            "parents": parents,
            "author": {"name": an, "email": ae, "date": ai},
            "committer": {"name": cn, "email": ce, "date": ci},
        })
        if code != 201:
            sys.exit("commit 创建失败 %s: %s %s" % (csha[:8], code, d))
        if d["sha"] != csha:
            sys.exit("commit sha 不一致！本地 %s vs 远端生成 %s（中止，未动 ref）" % (csha, d["sha"]))
        code, d = http("PATCH", "/git/refs/heads/" + cfg["branch"], token,
                       {"sha": csha, "force": False})
        if code != 200:
            sys.exit("ref 更新失败: %s %s" % (code, d))
        print("  ✓ %s  %s" % (csha[:7], msg_b.decode("utf-8", "replace").strip().splitlines()[0]))

    print("远端 HEAD 已到 %s" % local_sha[:12])


if __name__ == "__main__":
    main()
