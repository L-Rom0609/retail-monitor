#!/bin/bash
# macOS / Linux 启动器：拉云端 → 合并本地 CSV → 重建全部数据 → 推送上线
# 用法（在仓库根目录执行）：bash 脚本/update.sh [参数]
# 参数会透传给 脚本/full_sync.py，例如 --no-pull / --no-push / --push=api
set -e

HERE="$(cd "$(dirname "$0")/.." && pwd)"
cd "$HERE"

CFG="$HERE/配置与映射/sync_config.json"

# 首次运行：把环境变量里的令牌写进本地配置（该文件已被 .gitignore 忽略，不会入库）
if [ ! -f "$CFG" ] && [ -n "$RETAIL_SYNC_PAT" ]; then
  mkdir -p "$HERE/配置与映射"
  cat > "$CFG" <<EOF
{
  "remote": "origin",
  "branch": "main",
  "remote_url": "https://github.com/L-Rom0609/retail-monitor.git",
  "pat_env": "RETAIL_SYNC_PAT",
  "pat": "$RETAIL_SYNC_PAT"
}
EOF
  echo "已根据 RETAIL_SYNC_PAT 生成 配置与映射/sync_config.json（仅本机保存）"
fi

if [ ! -f "$CFG" ] || ! grep -q '"pat": *"[^"]' "$CFG"; then
  echo "提示：还没配置推送令牌。若系统 git 已存凭据可继续，否则只能本地生成。"
  echo "      推荐执行一次： export RETAIL_SYNC_PAT=你的GitHub令牌 && bash 脚本/update.sh"
  echo
fi

PY="${PYTHON:-python3}"
if ! command -v "$PY" >/dev/null 2>&1; then
  PY="python"
fi
if ! command -v "$PY" >/dev/null 2>&1; then
  echo "错误：找不到 Python 3。请先安装：brew install python3 或到 python.org 下载。"
  exit 1
fi
if ! command -v git >/dev/null 2>&1; then
  echo "错误：找不到 git。请先执行：xcode-select --install"
  exit 1
fi

"$PY" "$HERE/脚本/full_sync.py" "$@"
