# -*- coding: utf-8 -*-
"""
把 preview-ops.html + 三个 .bin 打包成一个零网络依赖的单文件 HTML。

用途：预览看板在「本地 file:// 双击」或「IDE 预览面板」里打开时，
相对路径 fetch 会被 CORS 拦截，CDN 回退在部分网络下也不可达 → 白屏。
本脚本把三个 gzip 数据包以 base64 内嵌进页面，window.__INLINE_DATA__ 优先命中，
页面不再发起任何网络请求。

用法： python 脚本/build_preview_standalone.py
输出： preview-ops-standalone.html
"""
import base64
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "preview-ops.html")
OUT = os.path.join(ROOT, "preview-ops-standalone.html")

BINS = [("dash", "dashboard-data.bin"), ("traf", "traffic-data.bin"), ("goods", "goods-data.bin")]

MARK = '<script id="inline-data">'


def main():
    if not os.path.exists(SRC):
        sys.exit("找不到 %s" % SRC)

    with open(SRC, "r", encoding="utf-8") as f:
        html = f.read()

    # 幂等：如果源文件里已经带内嵌块（有人误把产物当源），先剥掉
    html = re.sub(r'<script id="inline-data">.*?</script>\s*', "", html, flags=re.S)

    payload = []
    total = 0
    for key, name in BINS:
        path = os.path.join(ROOT, name)
        if not os.path.exists(path):
            sys.exit("缺少数据包 %s" % path)
        with open(path, "rb") as f:
            raw = f.read()
        total += len(raw)
        b64 = base64.b64encode(raw).decode("ascii")
        payload.append('  %s: "%s"' % (key, b64))
        print("  内嵌 %-18s gzip %6.0f KB → base64 %6.0f KB" % (name, len(raw) / 1024, len(b64) / 1024))

    block = (
        MARK + "\n"
        "/* 由 脚本/build_preview_standalone.py 生成，勿手改；源文件是 preview-ops.html */\n"
        "window.__INLINE_DATA__ = {\n" + ",\n".join(payload) + "\n};\n"
        "</script>\n"
    )

    # 插到 </head> 前，确保先于主脚本执行
    if "</head>" in html:
        html = html.replace("</head>", block + "</head>", 1)
    else:
        html = block + html

    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)

    print("\n输出 %s" % OUT)
    print("  数据合计 %.1f MB (gzip) → 文件 %.1f MB" % (total / 1024 / 1024, os.path.getsize(OUT) / 1024 / 1024))


if __name__ == "__main__":
    main()
