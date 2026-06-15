# 港股公告财经日历 — Agent 协作指南

## 项目一句话

根据 Excel 持仓自动爬取港股公告，生成可筛选的月度财经日历网页，部署在 GitHub Pages。

## 关键路径

| 文件 | 作用 |
|------|------|
| `src/crawler.py` | 唯一主程序：读取持仓 → 爬东方财富公告 → 输出 `index.html` / `logs/` / `data/` / `notices/` |
| `config/holdings.json` | 持仓配置缓存。优先从 Excel 读取，回退到此文件 |
| `index.html` | 财经日历页面，GitHub Pages 入口 |
| `auto_run.bat` | 定时任务脚本：爬取 + `git add/commit/push` |
| `setup_task.bat` | 创建 Windows 定时任务（需管理员权限） |
| `deploy.bat` | 手动首次部署 / 手动推送到 GitHub Pages |

## Agent 红线

1. **不要手动编辑 `index.html`**。它由 `crawler.py` 自动生成，下次运行会被覆盖。
2. **不要删除 `notices/` 或 `data/announcements.json` 中的历史数据**，除非用户明确要求清理磁盘。这些是公告缓存，GitHub Pages 展示需要它们。
3. **不要修改 `config/holdings.json` 里的持仓列表**（除非用户让改）。持仓应从 Excel 读取或让用户手动维护。
4. **日志和缓存不要进 git**。`logs/` 已被 `.gitignore` 忽略；`notices/` 和 `data/` 由 `auto_run.bat` 显式添加提交。

## 常见修改场景

### 133 招商局基金公告邮件通知

- 运行 `run.bat` 或 `auto_run.bat` 后，会自动调用 `scripts/notify_stock_133.py`
- 如果 `data/announcements.json` 中 `stock_code` 为 `00133` 且日期为当天的公告数量 > 0，则通过 Outlook 发邮件到 `ellenli@tfisec.com`
- 也可手动运行：`python scripts/notify_stock_133.py [--date YYYY-MM-DD]`

### 调整日历页面样式/筛选功能

改 `src/crawler.py` 里生成 HTML 的逻辑，然后运行：

```bash
python src/crawler.py
```

刷新浏览器查看 `index.html`。

### 修改持仓数据源

编辑 `config/holdings.json` 里的 `excel_path`，或传命令行参数：

```bash
python src/crawler.py --holdings "C:\path\to\持仓.xlsx"
```

### 部署到 GitHub Pages

- 首次：双击 `deploy.bat`，按提示粘贴仓库链接。
- 后续自动：运行 `setup_task.bat` 创建定时任务，或手动双击 `auto_run.bat`。

### Pages 健康检查 Loop

项目已实现一个最小 Loop：

- `scripts/check_pages_health.py`：检查 Pages 可访问性，异常时自动重推
- `check_pages.bat`：手动触发
- `scripts/setup_pages_check.bat`：创建每天 3:00 的定时任务（需管理员权限）

通知方式：异常时发邮件到 `ellenli@tfisec.com`（通过本机 Outlook），Outlook 不可用时降级为桌面通知。

**Loop 边界**：健康检查只做重推，不重新运行爬虫，不修改代码逻辑，不删除数据。

## 环境

- Python 3.8+
- 依赖：`requests`
- 安装：`pip install -r requirements.txt`
- 平台：Windows（ bat 脚本），但爬虫核心跨平台。
