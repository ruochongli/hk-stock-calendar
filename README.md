# 港股公告爬虫助手

每天根据持仓自动爬取对应港股上市公司的最新一周公告，格式化为简短摘要。

## 输出格式

```
【股票代码】【股票名称】【日期】发布公告【公告名称】
```

示例：
```
【00700】【腾讯控股】【2026-06-01】发布公告【翌日披露报表 - 已发行股份变动及股份购回】
【09988】【阿里巴巴-W】【2026-06-01】发布公告【公告 根据2024年计划授出奖励】
```

## 快速开始

### 1. 配置持仓（自动从 Excel 读取）

爬虫**默认从 Excel 持仓文件自动读取**，无需手动维护 JSON。

默认 Excel 路径：
```
C:\Users\ellenli\OneDrive - TFI\From 若冲\@@@@@2026@@@@@\June\CA每日监测文件 rc.xlsx
```

Excel 格式要求（第二个 sheet）：
| 基金名 | 账户名 | 股票代码 | 股票名称 |
|--------|--------|----------|----------|
| CDG | 50308215 | 257 | 光大環境 EB ENVIRONMENT |
| CDG | 50308215 | 700 | 騰訊控股 TENCENT |

如果 Excel 路径变了，可在 `config/holdings.json` 中配置：
```json
{
  "excel_path": "C:\\Users\\ellenli\\Desktop\\持仓.xlsx"
}
```

也可以回退到手动 JSON 配置（当 Excel 不存在时自动使用）：
```json
{
  "holdings": [
    {"code": "00700", "name": "腾讯控股", "fund": "基金A"},
    {"code": "09988", "name": "阿里巴巴-W", "fund": "基金B"}
  ]
}
```

### 2. 运行脚本

#### 方式一：双击运行（Windows）

直接双击 `run.bat`

运行完成后会自动检查 **133 招商局基金** 当日公告，如有则发邮件到 `ellenli@tfisec.com`。

#### 方式二：命令行运行

```bash
cd stock_crawler
python src/crawler.py
python scripts/notify_stock_133.py
```

### 3. 查看结果

运行后会在终端直接打印结果，同时自动生成两份文件：

| 文件 | 说明 |
|------|------|
| `logs/announcements_YYYYMMDD.txt` | 纯文本汇总 |
| `index.html` | **财经日历网页**，可直接用浏览器打开 |

**财经日历预览：**

- **完整月历视图**：6行 × 7列的标准日历网格，包含上个月和下个月的溢出日期
- **年月选择**：顶部下拉框直接选择任意年月
- **基金/账户筛选**：先选基金（如 CDG / DA / Lakeside），股票筛选栏自动联动只显示该基金下的股票
- **股票筛选**：勾选/取消勾选具体股票，精准过滤
- **今日高亮**：当前日期有蓝色边框标记
- **公告数量**：当天有公告时右上角蓝色角标显示数量
- **基金标签**：每条公告卡片顶部显示所属基金名，左侧带基金颜色边条
- **响应式**：桌面端、平板、手机自适应

打开方式：双击 `index.html` 或在浏览器中打开 `file:///C:/Users/ellenli/stock_crawler/index.html`

---

## 部署到网页（分享给同事）

如果你希望同事通过链接直接查看，而不是发文件，可以部署到 **GitHub Pages**（免费）：

**首次部署**：
1. 去 [github.com/new](https://github.com/new) 创建一个空仓库（如 `hk-stock-calendar`）
2. 双击 `deploy.bat`，按提示粘贴仓库链接
3. 去 GitHub 仓库 Settings → Pages → Source 选择 `main` 分支
4. 等待 1 分钟，获得 `https://你的用户名.github.io/hk-stock-calendar` 链接

**后续更新**：
- 爬虫运行后，双击 `deploy.bat` 即可自动同步到网页

详细步骤见 [`GITHUB_PAGES_SETUP.md`](GITHUB_PAGES_SETUP.md)。

## 设置每天自动运行（Windows 任务计划程序）

### 推荐：一键创建定时任务

直接双击 `setup_task.bat`（建议右键**以管理员身份运行**），会自动创建一个每天 9:00 运行的定时任务：

- 任务名：`HK-Stock-Calendar-Auto-Update`
- 执行内容：`auto_run.bat`
- 运行效果：自动爬取公告 → 推送到 GitHub Pages

`auto_run.bat` 会静默运行爬虫，并把 `index.html`、`notices/`、`data/announcements.json` 自动提交推送到 GitHub。

### 手动创建（备用）

1. 按 `Win + S`，搜索并打开 **"任务计划程序"**
2. 右侧点击 **"创建基本任务"**
3. **名称**：港股公告爬虫
4. **触发器**：选择 **"每天"**，设置运行时间（如早上 9:00）
5. **操作**：选择 **"启动程序"**
   - 程序或脚本：`C:\Users\ellenli\stock_crawler\auto_run.bat`
   - 起始于（可选）：`C:\Users\ellenli\stock_crawler`
6. 完成创建

> 提示：`auto_run.bat` 已设计为后台静默运行，无需删除 `pause`。

---

## GitHub Pages 健康检查 Loop

Pages 偶尔会因为网络、GitHub 延迟或推送失败出现 404。可以启用自动健康检查：

- **检查脚本**：`scripts/check_pages_health.py`
- **手动运行**：双击 `check_pages.bat`
- **创建定时任务**：右键以管理员身份运行 `scripts/setup_pages_check.bat`

默认行为：

- 每天 **3:00** 访问 `https://ruochongli.github.io/hk-stock-calendar/`
- 如果 HTTP 不是 200，自动 `git push` 重推现有 `index.html`
- 结果写入 `logs/pages_health_YYYYMMDD.log`
- 异常或修复失败时**发邮件到 ellenli@tfisec.com**（通过 Outlook；如果 Outlook 不可用，降级为 Windows 通知）

> 注意：健康检查只做"重推"，不会重新运行爬虫，因此不依赖 Excel 持仓文件。

## 项目结构

```
stock_crawler/
├── config/
│   └── holdings.json      # 持仓配置（自动生成/手动维护）
├── src/
│   └── crawler.py         # 主程序：读取持仓、爬取公告、生成日历
├── data/
│   └── announcements.json # 公告原始数据（自动生成）
├── notices/               # 公告详情 HTML 缓存（自动生成）
├── logs/                  # 运行日志（自动生成）
├── index.html             # 财经日历网页（自动生成，也是 GitHub Pages 入口）
├── run.bat                # 一键运行爬虫并生成本地日历
├── auto_run.bat           # 定时任务用：爬取 + 自动推送到 GitHub Pages
├── setup_task.bat         # 创建 Windows 定时任务（每天 9:00 自动更新）
├── deploy.bat             # 手动部署/首次部署到 GitHub Pages
├── GITHUB_PAGES_SETUP.md  # GitHub Pages 部署详细指南
├── requirements.txt       # Python 依赖
└── README.md              # 本文件
```

## 数据来源

- **东方财富** (`np-anotice-stock.eastmoney.com`)
- 数据延迟通常在 1 小时内，与港交所披露易保持同步

## 依赖

- Python 3.8+
- requests

安装依赖：
```bash
pip install -r requirements.txt
```
