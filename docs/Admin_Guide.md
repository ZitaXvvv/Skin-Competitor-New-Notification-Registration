# CI New SKU Tracker — 管理员操作手册

> 面向：负责服务器日常运维、数据纠错、账号管理的管理员。
> 技术栈/架构设计请见附件 [Architecture.md](Architecture.md)。

## 1. 系统总览

- 看板前端：`streamlit run src/dashboard.py`，默认监听 `8501` 端口，普通用户直接访问、无需登录。
- 数据源：`Documents\Olay CI\CI_List_Ada.xlsx`（每个品牌一个sheet），看板和抓取脚本读写的是同一份文件。
- 抓取流程：`src/main.py`，每月由 Windows 计划任务自动触发一次（见第2节）。
- 管理员登录后（看板左侧栏"🔒 管理模式"），会出现"🛠️ 管理面板"，可以手动触发抓取、查看上次运行状态、增/删/改单行数据。

# CI New SKU Tracker — 管理员操作手册

> 面向：负责服务器日常运维、数据纠错、账号管理的管理员。
> 技术栈/架构设计请见附件 [Architecture.md](Architecture.md)。

## 1. 系统总览

- 看板前端：`streamlit run src/dashboard.py`，默认监听 `8501` 端口，普通用户直接访问、无需登录。
- 数据源：`Documents\Olay CI\CI_List_Ada.xlsx`（每个品牌一个sheet），看板和抓取脚本读写的是同一份文件。
- 抓取流程：`src/main.py`，每月由 Windows 计划任务自动触发一次（见第2节）。
- **每日补全流程**：`src/module6_daily_fill.py`，每天凌晨3点自动补全10个不完整产品卡片（见第2.1节）。
- 管理员登录后（看板左侧栏"🔒 管理模式"），会出现"🛠️ 管理面板"，可以手动触发抓取、查看上次运行状态、增/删/改单行数据。

## 2. 计划任务配置（服务器首次部署必做）

### 2.0 一键注册所有计划任务（推荐）

以管理员身份打开 PowerShell，`cd` 到项目根目录，运行：

```powershell
.\scripts\setup_daily_fill_task.ps1
```

此脚本会同时注册以下两个任务：

| 任务名 | 触发时间 | 执行内容 |
|---|---|---|
| `CI_NewSKU_Monthly_Scrape` | 每月1号 02:00 | 全量抓取过去40天新品（`main.py --days 40`） |
| `CI_NewSKU_Daily_Fill` | 每天 03:00 | 增量补全10个不完整产品卡片（`module6_daily_fill.py`） |

注册后可在"任务计划程序"(`taskschd.msc`)查看和手动触发。

### 2.1 每日增量补全任务详解

**目标**：对所有历史年份的产品，以每天10个的速度补全缺失的产品卡片信息，优先顺序 2026→2025→2024。

**产品卡片完整定义**（以下任一缺失则该产品列为待补全）：
- 产品名称（通常抓取时已有）
- 备案号/注册号（通常抓取时已有）
- **成分列表**（`Ingredient` 列）
- **备案/注册 PDF 链接**（`link` 列）
- **产品图片**（`image_map.json` 中有记录）
- **mini-POC 链接**（`mini POC` 列，即 NMPA 产品详情页）

**每日任务执行顺序**：
1. 检查美丽修行（BEBD）Cookie 是否有效，记录到日志
2. 检查 NMPA hzpba 网站是否可访问，记录到日志
3. 扫描全部历史 Excel，找出不完整产品，按 2026→2025→2024 排序
4. 取前10个，逐一执行：
   a. 如有本地 PDF → 渲染首页为产品图（无需网络）
   b. 如是特殊注册 PDF（文字型）→ 提取全成分列表
   c. 如 BEBD/NMPA 可访问且缺少 PDF 链接 → 在线查询并写回 Excel
   d. 新查到 PDF 链接后自动下载 + 渲染产品图
5. **每7天运行一次历史月份补全**：
   - 扫描全部 Excel 中各月实际数据量，识别零数据月份
   - 对最近的缺失月份，自动计算所需 `--days` 参数并调用 `main.py` 补全
   - 一次补一个月，进度保存在 `log/backfill_state.json`
   - 若 BEBD 未登录则警告但不阻止执行，会以无登录状态尽力抓取
6. 写入日志 `log/daily_fill_YYYYMMDD.log`

**手动操作命令**：
```powershell
# 只检查连通性，不补全
python src\module6_daily_fill.py --check

# 预览会补全哪些产品和缺失月份（不实际写入）
python src\module6_daily_fill.py --dry-run

# 正常运行，每次10个 + 每7天一次历史月份补全
python src\module6_daily_fill.py

# 立即强制运行历史月份补全（不等7天间隔）
python src\module6_daily_fill.py --backfill

# 自定义每次处理数量
python src\module6_daily_fill.py --limit 20
```

### 2.2 每月全量抓取怎么检查
- 每月任务执行 `python src\main.py --days 40`（回溯40天，略多于一个月以防遗漏）。
- 检查/手动测试：`taskschd.msc` → 找到 `CI_NewSKU_Monthly_Scrape` → 右键"运行"。
- 也可以在看板"🛠️ 管理面板 → 📡 抓取任务"里直接点击"🚀 立即触发一次全量抓取"手动补跑（会显示上次运行的日志摘要和状态 ✅/❌）。

## 3. 什么时候需要登录什么网站（人工介入点）

这是最容易被忽略、也最容易导致自动化"卡住"的部分，请重点关注：

### 3.1 美丽修行 BEBD（bebd.bevol.com）— ⚠️ 需要人工定期刷新登录

**凌晨计划任务的处理机制（已解决卡死问题）**：
- 计划任务使用 `--unattended` 标志运行，Cookie 失效时**自动跳过 BEBD 抓取并记录警告日志**，不再阻塞等待。
- 跳过后日志会出现：`⚠️ BEBD 跳过（无人值守模式）: Cookie 已失效`

**如何刷新登录（工作时间操作，每月一次）**：
```powershell
.\scripts\refresh_bebd_login.ps1
```
脚本会打开有头浏览器 → 管理员完成登录 → 按 Enter 保存 Cookie → 当晚/当月凌晨任务自动复用。

**建议操作节奏**：
- **每月月底最后一个工作日下班前**执行一次登录刷新（Cookie 通常可维持数周）
- 看到每日日志中出现 `BEBD: ❌` 时立即执行刷新

**手动检查 Cookie 是否仍有效**：
```powershell
python src\module6_daily_fill.py --check
```

### 3.2 NMPA 政府网站（nmpa.gov.cn/datasearch、hzpba.nmpa.gov.cn）— 全自动，无需登录
- 用于查询"特殊化妆品注册"和"普通化妆品备案"的详情/PDF链接，全程自动化，不需要人工登录。
- ⚠️ **已知限制**：这两个网站从 2026-07-23 起对所有自动化访问方式（含正常浏览器请求）返回 400/412 空响应（阿里云WAF反爬），导致这部分数据暂时无法查到。这不是代码bug，请不要尝试绕过反爬机制。详见附件"已知限制"一节。

### 3.3 SharePoint / 邮件（Microsoft Graph API / Outlook）— 全自动，无需登录
- 上传 PDF 到 SharePoint、发送汇总邮件，都是用 Azure AD 应用程序（App Registration）的 Client Credentials 方式自动认证，不需要人工登录。
- 邮件发送优先走本机 Outlook（如果服务器装了 Outlook 且已登录企业邮箱），否则自动退回 Graph API 方式。
- ⚠️ **安全提醒**：当前 `src/config.py` 里的 Azure AD Client Secret 是明文硬编码，且已经提交进 GitHub 仓库历史记录。**强烈建议尽快在 Azure 门户里重置(rotate)这个 secret**，并改造成从环境变量或本地不进git的配置文件读取，避免旧 secret 被任何能访问该仓库历史的人利用。

### 3.4 GitHub 代码仓库 — 只有更新代码时才需要
- 仓库：`https://github.com/ZitaXvvv/Skin-Competitor-New-Notification-Registration`
- 用的是 SSH 部署密钥（本机 `~/.ssh/config` 里 `github-zitaxvvv` 这个 host 别名对应的密钥），日常运行/抓取不涉及这一步，只有需要拉取/推送代码更新时才用得到。

### 3.5 Windows Server 本身 — 服务器登录
- RDP/物理登录服务器使用的是服务器自身的 Windows 账号，与本系统无关，由 IT 常规管理。

## 4. 出错了怎么办

- **第一步永远是**：看板"🛠️ 管理面板 → 📡 抓取任务"会自动显示最近一次运行的日志尾部和 ✅/❌ 状态；或者直接去 `log/` 目录找最新的 `ci_bot_*.log` 文件看完整日志。
- **常见问题排查**：

| 症状 | 可能原因 | 处理方法 |
|---|---|---|
| 计划任务显示"运行中"但一直不结束 | 卡在等待 BEBD 人工登录（见3.1） | RDP登录服务器看终端提示；确认后手动登录一次刷新Cookie |
| 抓取数量为0或明显偏少 | BEBD Cookie 过期 / 网站改版 | 同上，手动跑一次 `--step 1` 验证 |
| 某些产品没有备案/注册链接、缺PDF | NMPA/hzpba网站被WAF拦截（已知限制，见3.2） | 无需处理，等网站策略变化；不要尝试绕过 |
| SharePoint上传/邮件发送失败 | Azure AD secret 失效或权限变更 | 检查 Azure 门户里该App Registration状态，必要时重置secret并更新config.py |
| Excel数据被误删/改错 | 管理面板操作或误触 | 面板本身每次编辑/删除都会先自动备份到 `Documents\Olay CI\_admin_backups\`，取最新时间戳的文件手动替换回`CI_List_Ada.xlsx`即可（替换前先关闭Excel和重启streamlit） |
| 看板打不开/端口占用 | streamlit进程未启动或崩溃 | RDP登录服务器，`cd`到项目目录后 `streamlit run src/dashboard.py` 重新启动（当前没有开机自启机制，服务器重启后需要手动这一步——如需要可以让开发者再加一个开机自启的计划任务） |

## 5. 管理员账号管理

- 新增管理员：RDP登录服务器 → `cd` 到项目目录 → 运行：
  ```
  python src/manage_admins.py add <用户名>
  ```
  按提示输入两次密码（不回显）。**请管理员自己在终端里输入真实密码，不要把密码告诉他人代为设置。**
- 删除管理员：`python src/manage_admins.py remove <用户名>`
- 查看现有管理员列表：`python src/manage_admins.py list`
- 每个管理员密码独立、加盐哈希存储在 `src/admins.json`（不进git），互相看不到明文密码。
- 所有增/删/改操作都会记录到 `log/admin_actions.log`（谁、什么时候、改了什么），便于追溯。

## 6. 每月人工检查清单

- [ ] 月度任务（`CI_NewSKU_Monthly_Scrape`）当月是否按时运行，日志末尾无 ERROR
- [ ] 每日补全任务（`CI_NewSKU_Daily_Fill`）是否每天正常触发（看 `log/daily_fill_*.log`）
- [ ] BEBD 登录 Cookie 状态：看每日日志里的 `BEBD: ✅/❌` 标记；若连续出现 ❌，需手动重登
- [ ] NMPA hzpba 连通性：看每日日志里的 `NMPA: ✅/❌`；❌ 属已知 WAF 限制，不影响本地图片渲染
- [ ] 待补全产品数量趋势：运行 `python src\module6_daily_fill.py --dry-run` 查看"剩余未完整产品"数
- [ ] SharePoint 上传 / 邮件发送这两步日志末尾是否显示成功
- [ ] （若近期没做过）确认 Azure AD Client Secret 是否已经按第3.3节建议重置

## 附件

详细技术栈、模块设计、数据流、设计取舍与已知限制，见 [Architecture.md](Architecture.md)。
