# CI New SKU Tracker — 使用手册 / Quick User Guide

## 这是什么？ What is this?
一个网页看板，按季度月历展示各竞品品牌的新备案/注册产品（成分、功效宣称、Artwork PDF、功效证明、产品图）。
A web dashboard showing competitor brands' newly filed/registered products by quarter (ingredients, benefit claims, artwork PDF, proof-of-claim, product image).

**覆盖品牌**：珀莱雅、谷雨、欧诗漫、兰蔻、欧莱雅、雅诗兰黛、修丽可、百雀羚、韩束、自然堂、薇诺娜、科颜氏、娇韵诗。

## 打开方式 How to open
向管理员获取访问地址（形如 `http://<服务器地址>:8501`），浏览器打开即可，**无需登录**。
Ask your admin for the URL (e.g. `http://<server>:8501`) and open it in a browser — **no login required** for browsing.

## 筛选栏 Filters (top of page)
| 控件 Control | 作用 What it does |
|---|---|
| 品牌 Brand | 多选要看的品牌，默认全选 / multi-select brands, all selected by default |
| Year（最多3个 max 3） | 选择要显示的年份 / select which year(s) to show |
| Type | 全部 / 普通备案 Filing / 特殊注册 Registration |
| ✔ All | 一键恢复全部品牌 / reset to all brands |
| EN 开关 toggle | 打开后产品卡片名称切换为英文翻译 / toggle to show English product name on each card |

## 统计卡片 Summary cards
页面顶部显示当前筛选下的汇总数字：
- **Total Products** — 产品总数
- **🟢 普通备案** — 普通化妆品备案数量
- **🔴 特殊注册** — 特殊化妆品注册数量
  > ⚠️ 特殊化妆品在注册提交时不需要提交 artwork，所以只有标签样稿（无产品包装图）
- **Active Brands** — 当前选中品牌数

## 看懂卡片 Reading a product card
- **蓝色边框** = 普通备案（Filing）
- **红色边框** = 特殊注册（Registration）
- **产品图** = 有图时显示产品包装照或备案 PDF 首页截图；无图时显示灰色瓶型占位符（特殊注册产品通常无图）

卡片右下角图标：
- 🖼 = Artwork / 备案注册 PDF（点击在新标签页打开）
- 🏷 = 产品标签 Label 样稿
- 🧪 = NMPA 官网产品详情 / mini-POC 功效证明

## 成分对比 Comparing ingredients
把卡片**拖拽**到页面悬浮的对比框里，可以并排比较多个产品的全成分。
Drag a card into the floating comparison panel to compare full ingredient lists side by side.

## 下载 PDF Downloading PDFs
页面下方"PDF 下载区"可直接一键下载本地已缓存的 PDF（公司网络内打不开外部 NMPA 链接时优先用这个）。
Use the "PDF Download" section further down the page to download locally cached PDFs directly (useful if the external NMPA link is blocked on the company network).

## 原始数据 Raw data
底部"📊 原始数据"可展开查看表格形式的完整数据。
Expand "📊 原始数据" at the bottom to see the full data as a table.

## 数据多久更新一次？ How often is data refreshed?
- **每月1号凌晨**：自动抓取过去40天的新品数据
- **每天凌晨**：自动补全缺失的产品图片、成分、PDF链接（每次约10个产品）
- **历史补全**：系统会每周自动回补历史缺失月份的数据

发现数据有误，请联系管理员在后台管理面板修正，不要自行编辑源文件。
If you spot an error, contact an admin to correct it via the admin panel — do not edit source files yourself.
