# Mac 端部署与双机协同方案

> 目标：① 这套看板更新任务能在 Mac 上正常跑；② Mac 也能作为「上传数据的设备」，把新底表导入后直接推送上线。
> 结论：**可以，且不需要装任何第三方 Python 包**——全项目只用 Python 标准库，Mac 上只需 Python 3 + git + 一枚 GitHub 令牌。

---

## 一、先说结论与本次已完成的改造

项目原本就宣称「Windows / macOS 双端对等」，但实际有 5 处会让 Mac 直接踩坑（启动器路径写错、配置文件写到错误目录、脚本无可执行位、README 里的输入目录名不存在、双击入口缺失）。本次已全部修好并推到仓库，Mac `git clone` 后即可使用：

| 文件 | 改动 |
|---|---|
| `脚本/full_sync.py` | **新增**。把过去需要人工分步做的「抽 bin → 同步 index → 重建流量/商品 → 钉 CDN → 推送」代码化成一条命令，两端通用 |
| `脚本/update.sh` | **修复**。原先 `cd 脚本/` 导致配置写到 `脚本/sync_config.json`（代码读的是 `配置与映射/`），且调用路径错误；现统一以仓库根为基准 |
| `Mac 一键更新.command` | **新增**。Finder 双击即可跑完整流程，不用开终端 |
| `.gitattributes` | **新增**。统一 `.sh`/`.command` 为 LF，二进制与 GBK 的 `.bat` 不做转码 |
| `README.md` | 修正 Mac 章节里不存在的 `底表输入/`、`./update.sh` 路径 |

---

## 二、架构选择（三选一，推荐 A）

| 方案 | 做法 | 优点 | 缺点 |
|---|---|---|---|
| **A. 双机对等 + 串行（推荐）** | 两台都 clone 同一仓库，各自放 CSV、各自跑 `update.sh`，每次自动先拉云端再合并 | 任一台都能独立上线；历史存档随仓库同步 | 两端**不能同时跑**；同一天数据不要在两台重复导入 |
| B. Mac 只上传、Windows 统一处理 | Mac 把 CSV 放到共享云盘/私有仓库，Windows 定时拉取并统一跑流水线 | 零冲突、口径最稳 | Mac 不能独立完成上线 |
| C. 云端跑（GitHub Actions） | 只把 CSV 上传仓库，Actions 自动重建并发布 | 无需本地环境 | 要改造流水线 + 处理大文件与密钥，工程量大 |

**推荐 A**，配合下方「双机铁律」即可安全。若你们日常是「Windows 白天导、Mac 晚上导」，A 完全够用。

---

## 三、Mac 一次性安装（约 10 分钟）

```bash
# 1) 依赖：Python 3 + git（macOS 不自带，二选一）
xcode-select --install                 # 提供 git
brew install python3                   # 或到 python.org 下载 pkg 安装
python3 --version && git --version      # 确认都能输出版本号

# 2) 克隆（务必用 git clone，不要用 AirDrop / U 盘拷贝目录）
git clone https://github.com/L-Rom0609/retail-monitor.git
cd retail-monitor

# 3) 中文路径保护 + 换行符（一次性）
git config core.autocrlf input
git config core.precomposeunicode true   # 关键：避免 macOS NFD 与仓库 NFC 中文目录名不匹配
chmod +x 脚本/update.sh "Mac 一键更新.command"

# 4) 令牌：去 GitHub → Settings → Developer settings → Personal access tokens
#    建一枚 fine-grained token，仅给该仓库的 Contents: read/write
export RETAIL_SYNC_PAT=你的令牌
echo 'export RETAIL_SYNC_PAT=你的令牌' >> ~/.zshrc   # 以后免手工输入

# 5) 首次跑一次（会自动把令牌写进 配置与映射/sync_config.json，该文件不入库）
bash 脚本/update.sh
```

> **务必用 `git clone`**。中文目录名（`配置与映射/`、`数据底表/`、`门店明细/`）在 macOS 上以 NFD 存储、仓库里是 NFC；手工拷贝或 Finder 新建极易造成 `FileNotFoundError`。`core.precomposeunicode true` + clone 可以规避。

---

## 四、日常「上传数据」怎么操作

### 放文件（三个约定位置）

| 数据 | 放到 | 说明 |
|---|---|---|
| 订单 CSV | `数据底表/` | 文件名形如 `订单数据2026-11-09-xxxx.csv`；也可设 `RETAIL_DATA_DIR=~/我的底表` |
| 门店明细 | `数据底表/门店明细/` | 文件名含「全部门店-门店明细」，流量页靠它 |
| 归属表 / 成本表 / 目标 | `配置与映射/` | **随仓库同步，不要手工改**；归属表唯一真源就是 `配置与映射/仓名-负责人归属关系表.xlsx` |

### 跑更新（二选一）

```bash
# 终端方式（推荐，能看到全过程日志）
bash 脚本/update.sh

# 或 Finder 双击
# 「Mac 一键更新.command」
```

它会依次：拉云端 → 合并本机 CSV 到历史存档 → 抽 `dashboard-data.bin` 并同步 `index.html` → 重建流量页与商品数据 → 提交 → 钉 CDN 版本 → 推送 → 打印线上地址。

常用参数：

```bash
bash 脚本/update.sh --no-pull      # 断网时只做本地重建（注意：双机场景别用）
bash 脚本/update.sh --no-push      # 只生成不推送，先本地检查
bash 脚本/update.sh --push=api     # github.com:443 被拦时，强制走 REST API 推送
RETAIL_DATA_DIR=~/Downloads/底表 bash 脚本/update.sh
```

### 验收

打开 https://l-rom0609.github.io/retail-monitor/ ，看左上角「数据截至」日期是否为刚导入的那天；若没变，等 1~2 分钟（Pages 构建）再刷新。

---

## 五、双机协同铁律（**最重要，决定数据会不会丢**）

1. **同一时间只跑一台**。两端同时推送会造成分叉与覆盖。养成习惯：跑之前先确认另一台没在跑。
2. **每次都先拉云端**。`update.sh` 默认会 `pull --rebase`；**不要加 `--no-pull` 上传数据**。
3. **同一天的 CSV 只在一台导入**。若两端都导入同一天，后跑者按 `日期|小时|仓名|平台` 覆盖前者。
4. **不要跑 `python 脚本/sync.py --squash`**。历史压缩会强推重写提交链，另一台立刻失去共同祖先，且已钉的 CDN `@sha` 会 404。需要压缩时，压缩完两端都重新 clone 一次。
5. **归属表只改仓库里那一份**，改完提交；另一台 `git pull` 后再跑。归属表旧版会把「表里已不存在的仓」历史物理删除（历史上出过删掉 30 仓的事故）。
6. **原始 CSV 不入库**（`.gitignore` 忽略了 `数据底表/`），推送的只是聚合后的 `历史数据/历史数据存档.json`。所以：
   - 日常增量更新没问题——Mac 拉到的存档已含全部历史；
   - 但 **Mac 无法独立做「全量重算 / --replace」**，因为缺 7~10 月原始订单。若需要这个能力，见下一节。

---

## 六、让 Mac 也具备「全量重算」能力（可选）

三选一，按你们的实际情况：

| 方式 | 做法 | 适合 |
|---|---|---|
| **共享云盘**（最简单） | 把历史原始 CSV 放 iCloud/企业微信盘/OneDrive，Mac 端 `RETAIL_DATA_DIR=~/云盘/历史底表 bash 脚本/update.sh --no-push` 后人工核对 | 偶尔重算 |
| **单独私有仓库** | 建 `retail-monitor-raw` 私有仓库存 CSV，两端各自 clone；重算时 staging 到 `.tmp_rebuild_input/` 再 `--replace` | 经常重算、要可追溯 |
| **定期拷贝** | 每月把 Windows 的历史底表拷到 Mac 的 `数据底表/历史底表/`（脚本不递归扫描该子目录，需手工 staging） | 最低成本 |

> 注：`update_dashboard.py` 只扫描 `数据底表/` 顶层（不递归），历史 CSV 不要直接堆进去影响日常增量。

---

## 七、安全与故障排查

**令牌**
- 当前 Windows 上的 `配置与映射/sync_config.json` 是明文 PAT 且被 git 忽略，Mac clone 不到，必须自己配一枚。
- 建议**轮换**：吊销旧令牌，新建 fine-grained token（仅该仓库 `Contents: read/write`），两端各存一份，不要写进任何会提交的文件。

| 现象 | 原因 / 处理 |
|---|---|
| `Permission denied: 脚本/update.sh` | `chmod +x 脚本/update.sh "Mac 一键更新.command"`，或直接用 `bash 脚本/update.sh` |
| `FileNotFoundError: 配置与映射/...` | 中文目录 NFD/NFC 不匹配 → 重新 `git clone`，并设 `core.precomposeunicode true` |
| 没找到订单底表 | CSV 放错目录。正确位置是 `数据底表/`（README 旧版写 `底表输入/` 是错的，已修） |
| `Tunnel connection failed: 502` | 本机代理问题。`full_sync.py` 走 REST 推送时会自动绕开代理；Mac 若必须走代理，设 `RETAIL_KEEP_PROXY=1` |
| `github.com:443` 被拦 | `bash 脚本/update.sh --push=api`（REST API 走 `api.github.com`） |
| `pull --rebase` 冲突 | 多为两端并发。先 `git status`，把冲突文件交维护者合并；**不要** `git reset --hard` |
| 线上数字没变 | Pages/CDN 延迟 1~2 分钟；确认页面里 `@<sha>/dashboard-data.bin` 已是新提交号 |

---

## 八、一句话流程

> Mac 把新 CSV 丢进 `数据底表/`（门店明细丢进 `数据底表/门店明细/`）→ 双击 `Mac 一键更新.command` → 等它打印「线上已刷新」→ 打开链接核对日期。
> Windows 端继续用 `一键更新看板.bat`，两端不要同时跑。
