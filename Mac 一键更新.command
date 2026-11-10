#!/bin/bash
# Finder 双击即可运行：把订单 CSV 放进 数据底表/ 后双击本文件
cd "$(dirname "$0")" || exit 1
echo "=============================================="
echo "  即时零售看板 · Mac 一键更新"
echo "  目录：$(pwd)"
echo "=============================================="
bash 脚本/update.sh "$@"
EC=$?
echo
if [ "$EC" -eq 0 ]; then
  echo "[完成] 线上已刷新：https://l-rom0609.github.io/retail-monitor/"
else
  echo "[失败] 退出码 $EC，请把上面的报错信息发给维护者。"
fi
echo
read -r -p "按回车键关闭窗口…" _
