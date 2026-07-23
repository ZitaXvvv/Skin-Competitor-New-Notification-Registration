# CI New SKU Tracker — 管理员操作手册

> 面向：负责服务器日常运维、数据纠错、账号管理的管理员。
> 技术栈/架构设计请见附件 [Architecture.md](Architecture.md)。

## 1. 系统总览

- 看板前端：`streamlit run src/dashboard.py`，默认监听 `8501` 端口，普通用户直接访问、无需登录。
- 数据源：`Documents\Olay CI\CI_List_Ada.xlsx`（每个品牌一个sheet），看板和抓取脚本读写的是同一份文件。
- 抓取流程：`src/main.py`，每月由 Windows 计划任务自动触发一次（见第2节）。
- 管理员登录后（看板左侧栏"🔒 管理模式"），会出现"🛠️ 管理面板"，可以手动触发抓取、查看上次运行状态、增/删/改单行数据。

## 2. 每月自动运行怎么配置 / 怎么检查

- 一次性设置（新服务器/首次部署时执行）：以管理员身份打开 PowerShell，`cd` 到项目根目录，运行：
  ```powershell
  .\scripts\setup_monthly_task.ps1
  ```
  默认注册计划任务 `CI_NewSKU_Monthly_Scrape`：每月1号凌晨2点自动执行 `python src\main.py --days 31`。
- 检查/手动测试：打开"任务计划程序"(`taskschd.msc`)，找到该任务，右键"运行"可以立即手动触发一次，用于验证配置是否正确。
- 也可以在看板"🛠️ 管理面板 → 📡 抓取任务"里直接点击"🚀 立即触发一次全量抓取"手动补跑（会显示上次运行的日志摘要和状态 ✅/❌）。

## 3. 什么时候需要登录什么网站（人工介入点）

这是最容易被忽略、也最容易导致自动化"卡住"的部分，请重点关注：

### 3.1 美丽修行 BEBD（bebd.bevol.com）— ⚠️ 需要人工登录
- `module1_bebd.py` 用 Playwright 打开一个**可见**的浏览器窗口访问 bebd.bevol.com。
- **首次部署，或者登录 Cookie 过期时**，脚本会在终端打印提示并暂停，等待人工在弹出的浏览器窗口里用账号密码或扫码完成登录，然后回到终端按 Enter，脚本才会继续（登录状态会存入 `src/bebd_cookies.json`，之后自动复用，不用每次都登录）。
- **风险**：如果这一步发生在无人值守的凌晨2点计划任务里，脚本会一直卡在等待终端输入，直到任务超时（`setup_monthly_task.ps1` 里设置的执行时限是3小时，超时后任务会被强制结束，本次抓取视为失败）。
- **建议**：每月计划任务跑完后，检查一次运行状态（第4节）；一旦发现失败或抓取数量异常，登录服务器手动跑一次 `python src/main.py --step 1`（这次会用可见浏览器），确认能正常自动登录（说明 Cookie 仍有效）或需要重新手动登录一次刷新 Cookie。

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

- [ ] 计划任务当月是否按时运行（`taskschd.msc` 查看"上次运行时间/结果"，或看板管理面板里的状态）
- [ ] 抓取数量是否明显异常（0条 或 暴增）
- [ ] BEBD 登录 Cookie 是否仍然有效（看日志有没有卡在登录提示）
- [ ] SharePoint 上传 / 邮件发送这两步日志末尾是否显示成功
- [ ] （若近期没做过）确认 Azure AD Client Secret 是否已经按第3.3节建议重置

## 附件

详细技术栈、模块设计、数据流、设计取舍与已知限制，见 [Architecture.md](Architecture.md)。
