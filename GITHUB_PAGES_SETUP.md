# GitHub Pages 部署指南

把港股公告日历部署到 GitHub Pages，这样同事只需一个链接就能查看，无需安装任何软件。

## 第一次部署（只需一次）

### 步骤 1：在 GitHub 创建仓库

1. 打开 [github.com/new](https://github.com/new)
2. **Repository name** 填写：`hk-stock-calendar`（也可以自定义）
3. **Visibility** 选择：
   - `Public`（公开，免费，任何人可访问）
   - `Private`（私密，GitHub Pro 才能开启 Pages）
4. **不要勾选** "Add a README file"
5. 点击 **Create repository**
6. 复制页面上的 HTTPS 链接，格式如：
   ```
   https://github.com/你的用户名/hk-stock-calendar.git
   ```

### 步骤 2：运行部署脚本

双击 `deploy.bat`，按提示粘贴刚才复制的仓库链接。

脚本会自动：
- 关联远程仓库
- 推送 `index.html`
- 显示你的 Pages 链接

### 步骤 3：开启 GitHub Pages

1. 打开你的 GitHub 仓库页面
2. 点击顶部 **Settings** → 左侧 **Pages**
3. **Source** 选择：
   - Branch: `main`
   - Folder: `/(root)`
4. 点击 **Save**
5. 等待约 1 分钟
6. 页面会显示访问链接，例如：
   ```
   https://你的用户名.github.io/hk-stock-calendar
   ```

### 步骤 4：分享给同事

把上面的链接发到微信/钉钉/邮件，同事点击即可查看。

---

## 后续更新（每天运行）

爬虫每天运行后，`index.html` 会自动更新。此时只需：

**双击 `deploy.bat`** → 自动推送到 GitHub → 等待 30 秒 → 网页自动刷新

或者把 deploy 加入 `run.bat` 的末尾，实现一键爬取+部署。

---

## 常见问题

### Q: 推送时提示需要输入用户名密码？

GitHub 已不支持密码登录，请使用 **Personal Access Token**：

1. 访问 [github.com/settings/tokens](https://github.com/settings/tokens)
2. 点击 **Generate new token (classic)**
3. 勾选 `repo` 权限
4. 生成后复制 token（只显示一次）
5. 推送时密码栏粘贴这个 token 代替密码

### Q: 页面打开是 404？

首次部署后需要等待 1-3 分钟。如果超过 5 分钟：
- 检查 Settings → Pages 中 Source 是否已选择 `main` 分支
- 确认仓库根目录有 `index.html` 文件

### Q: 如何更换域名？

在仓库根目录新建 `CNAME` 文件，内容写你的域名（如 `calendar.yourcompany.com`），然后在域名 DNS 添加 CNAME 记录指向 `你的用户名.github.io`。

### Q: 国内访问慢？

GitHub Pages 在国内偶有波动，如需稳定快速访问，建议改用：
- 腾讯云 COS + 静态网站托管
- 阿里云 OSS + 静态网站托管
- Vercel / Netlify（海外，但速度通常不错）
