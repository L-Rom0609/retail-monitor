#!/bin/bash
# Mac 端启动器：双端同步更新（拉云端 → 合并本地 CSV → 生成看板 → 推送，线上自动刷新）
# 用法：./update.sh  [其它传给 update_dashboard.py 的参数]
set -e
cd "$(dirname "$0")"

# 首次运行：把环境变量里的令牌写进本地 sync_config.json（该文件不入库）
if [ ! -f sync_config.json ] && [ -n "$RETAIL_SYNC_PAT" ]; then
  cat > sync_config.json <<EOF
{
  "remote": "origin",
  "branch": "main",
  "remote_url": "https://github.com/L-Rom0609/retail-monitor.git",
  "pat_env": "RETAIL_SYNC_PAT",
  "pat": "$RETAIL_SYNC_PAT"
}
EOF
  echo "已根据 RETAIL_SYNC_PAT 生成 sync_config.json（仅本机）"
fi

if [ ! -f sync_config.json ] || ! grep -q '"pat": *"[^"]' sync_config.json; then
  echo "提示：还没有配置推送令牌，本次只能本地更新、无法同步到线上。"
  echo "      执行一次即可： export RETAIL_SYNC_PAT=你的GitHub令牌 && ./update.sh"
  echo
fi

PY="${PYTHON:-python3}"
if ! command -v "$PY" >/dev/null 2>&1; then
  PY="python"
fi

"$PY" sync.py "$@"
