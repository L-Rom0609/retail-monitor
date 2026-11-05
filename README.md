# 即时零售看板 · 双端同步仓库（Windows + macOS 对等）

把 `retail-monitor` 看板的历史数据仓库与归属表放进 **Git 仓库**，
让 **Windows 和 Mac 都成为「对等的上传/更新节点」**：任一端处理完新底表，
另一端再处理时不会丢数据，最终发布的看板是两边合并后的完整结果。
（看板通过 GitHub Pages 托管，故仓库需为**公开**仓库；数据本就通过分享链接公开，无新增暴露。）

---

## 一、这个仓库里有什么

| 文件 | 作用 | 是否入库 |
|---|---|---|
| `update_dashboard.py` | 看板数据处理核心脚本（仅标准库，跨平台） | ✅ 是 |
| `sync.py` | Git 同步包装器：pull → 处理 → commit → push（冲突自动重试） | ✅ 是 |
| `历史数据存档.json` | 全部历史数据本体（唯一真源，按 `日期\|小时\|仓名\|平台` 聚合） | ✅ 是 |
| `仓名-负责人归属关系表.xlsx` | 门店→仓名→负责人映射（双端共享真源） | ✅ 是 |
| `月度目标配置.json` | 负责人月度目标 | ✅ 是 |
| `retail-dashboard.html` | 看板页面源文件（启动先走 CDN 竞速，失败回退同源 `dashboard-data.bin`，再回退内嵌） | ✅ 是 |
| `index.html` | 看板部署入口（= retail-dashboard.html 的副本，GitHub Pages 根文件） | ✅ 是 |
| `dashboard-data.bin` | 看板外部数据（列式+gzip 二进制，比旧版小 53%）；线上看板从 CDN 优先拉取它 | ✅ 是 |
| `data.json` / `dashboard-data.json` | 中间产物 | ❌ 不入库 |
| `底表输入/` | 你放 CSV / 归属表的本地目录 | ❌ 不入库（仅本机） |
| `sync_config.json` | 你的远程地址 / 令牌（个人） | ❌ 不入库 |

---

## 二、搭建状态（已完成，仅供留档）

| 项目 | 值 |
|---|---|
| GitHub 仓库 | https://github.com/L-Rom0609/retail-monitor （公开） |
| 线上看板 | **https://l-rom0609.github.io/retail-monitor/** |
| 数据文件 | https://gcore.jsdelivr.net/gh/L-Rom0609/retail-monitor@main/dashboard-data.bin |
| 托管方式 | GitHub Pages，来源 `main` 分支根目录（push 后自动重建） |
| 令牌保存 | 本机 `sync_config.json` 的 `pat` 字段（已被 `.gitignore` 忽略，不入库） |

> 若换机器（例如 Mac），只需：`git clone` 仓库 → 在同目录建一份 `sync_config.json`
> （复制 `sync_config.example.json`，把 `pat` 填成令牌）→ 即可运行。`remote_url` 已在模板里写好。

---

## 三、Windows 端日常使用

1. 把新的 **订单 CSV** 放进桌面 `数据看板--底表更新`（和以前一样）。
2. 要加新门店时，编辑仓库里的 **`仓名-负责人归属关系表.xlsx`**（不再是桌面那份）。
3. 双击 `一键更新看板.bat`。
   它会：拉最新 → 处理本地 CSV → 合并进历史存档 → 重新生成 `index.html` + `dashboard-data.bin`
   → 自动 push 到 GitHub → **线上看板几十秒后自动变成最新**（无需任何发布操作）。
4. 更新日志写在 `数据看板--底表更新/更新日志.txt`。

> 也可以把 CSV 直接丢进仓库里的 `底表输入/` 目录 —— 两个位置都会被扫描。

---

## 四、Mac 端安装与日常使用

```bash
# 1) 安装依赖（一次性）
xcode-select --install        # 提供 git
brew install python3          # 或用官网 pkg

# 2) 克隆仓库
git clone https://github.com/L-Rom0609/retail-monitor.git
cd retail-monitor

# 3) 配置令牌（一次性；之后会写进本地 sync_config.json，不再需要 export）
cp sync_config.example.json sync_config.json
# 编辑 sync_config.json，把 "pat": "" 改成 "pat": "你的令牌"
# 或：export RETAIL_SYNC_PAT=你的令牌   # update.sh 会自动帮你生成配置

# 4) 日常更新：把订单 CSV 丢进仓库里的 底表输入/ 目录，然后
./update.sh
# 或者让脚本去指定目录找 CSV：
RETAIL_DATA_DIR=~/我的底表 ./update.sh
```

Mac 与 Windows 行为完全一致：拉取云端最新 → 处理本地 CSV → 合并 → 推送 → 生成本地看板。

---

## 五、关键约定（避免两台机器互相覆盖）

- **归属表唯一真源 = 仓库里的 `仓名-负责人归属关系表.xlsx`**。两端都改这份、跑同步自动上传；
  不要再改桌面旧副本。
- **历史存档按聚合键合并**：`日期|小时|仓名|平台`，同键覆盖、异键追加 —— 两端处理不同日期的表是并集，互不丢失。
- **谁先推送谁赢**：若两端几乎同时跑，后推送的一方会被拒绝，脚本自动 `rebase` 重新合并后再推，
  无需人工处理。
- **原始 CSV 不入库**：只同步聚合后的 `历史数据存档.json`，仓库体积可控。

---

## 六、发布 / 更新线上看板（GitHub Pages，自动刷新）

**为什么用 GitHub Pages 而不是 WorkBuddy 发布**：WorkBuddy 发布平台会冻结一个应用的
静态文件，且本会话内复用同一应用——重新发布不会刷新线上文件（已实测：重发布后线上仍是
10-31 旧数据）。而 GitHub Pages 由仓库内容驱动：**push 即更新**，看板永远读最新。

**原理**：`index.html` 启动时先从 CDN（gcore/cdn.jsdelivr 并行竞速）拉 `dashboard-data.bin`，
都失败才回退同源与内嵌数据。所以「更新数据」= 跑 `sync.py` 让它把新 `dashboard-data.bin` **push** 到
仓库，GitHub Pages 在几十秒到几分钟内自动刷新线上看板，**无需任何重新发布操作**。

### 首次发布（✅ 已完成）
1. GitHub 公开仓库已建：`L-Rom0609/retail-monitor`。
2. 仓库已推送，GitHub Pages 已开启（来源：`main` 分支根目录）。
3. 线上地址：**https://l-rom0609.github.io/retail-monitor/**
4. 旧的 WorkBuddy 链接（`?sharecode=W3VCYBIKZB`）已冻结作废，可忽略（它永远停在 10-31 那版）。

### 之后每次更新（你只做这一件）
在 Windows 双击 `一键更新看板.bat`（或 Mac 跑 `./update.sh`）→ 脚本处理 CSV、合并、
生成 `dashboard-data.bin` 并自动 push（还会自动清 CDN 缓存） → **线上看板自动变成最新**，你无需任何发布操作。

> 本地 `retail-dashboard.html` / `index.html` 仍带内嵌数据（当前 148,809 行 / 截止 11-01），
> 断网或 GitHub 打不开时会自动回退到内嵌数据，双击也能在浏览器看。

---

## 七、仓库体积：历史自动压缩（无需你操心）

**为什么会变大**：每次更新都会新增一版 `历史数据存档.json`（16 MB 原文）+ 页面 + 数据文件，
Git 压缩存起来后**每次约 +2~3 MB**。

**自动策略**：`sync.py` 每次推送成功后自检，满足任一条件就自动压缩历史（把历史压成 1 条提交）：

| 触发条件 | 默认阈值 | 大致相当于 |
|---|---|---|
| 提交数 | ≥ 60 条 | 每天更新一次 ≈ 每 2 个月 |
| `.git` 体积 | ≥ 200 MB | 体积涨太快时提前触发 |

压缩动作 = 以当前文件快照重建一条提交 → force push → 本地 `gc` 清掉旧快照。
实测：12 条提交 / 29 MB → **1 条提交 / 2.7 MB**，所有数据文件完好。

**另一台机器怎么办**：它的下次运行会发现本地与云端没有共同祖先，
`git_pull` 会自动 `fetch + reset --hard` 对齐云端（日志会提示「已重置为云端最新」）。
原始 CSV 仍躺在底表目录，重跑流水线会重新合并，**不会丢数据**。

**手动控制**：
```bash
python sync.py --squash      # 立刻压缩一次（不等阈值）
python sync.py --no-squash   # 本次跳过我自动压缩
```
阈值可临时改：`RETAIL_SQUASH_MIN_COMMITS=30`、`RETAIL_SQUASH_MIN_MB=100`。

> 所以**不需要定闹钟提醒**——脚本自己盯着体积，到点就压。

---

## 八、常见问题

- **推送报 403 / 鉴权失败**：检查 `sync_config.json` 里的 `pat` 是否正确、令牌是否勾了 `repo`、是否已过期。
- **线上还显示旧数据**：Pages 重建需要几十秒，等 1~2 分钟刷新；仍不对就看仓库 Actions/Pages 是否构建失败。
- **想从头重算**：`python update_dashboard.py --replace`（会清空历史存档，慎用）。
- **Mac 找不到 git**：先 `xcode-select --install`。
