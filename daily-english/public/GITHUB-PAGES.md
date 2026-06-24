# GitHub Pages 发布说明

这个目录已经按 GitHub Pages 发布方式准备好了。

## 你要做的事

1. 把当前仓库推到 GitHub。
2. 打开仓库页面。
3. 进入 `Settings` -> `Pages`。
4. 在 `Build and deployment` 里选择 `GitHub Actions`。
5. 确保默认分支是 `main` 或 `master`。
6. 之后每次你把 `daily-english/public/` 的更新推上去，GitHub 会自动发布。

## 发布后的链接

发布成功后，网页通常会出现在：

- `https://你的用户名.github.io/你的仓库名/`

首页会自动打开：

- `index.html`

今天这页的固定地址会是：

- `https://你的用户名.github.io/你的仓库名/2026-06-24.html`

## 你需要分享什么

分享上面的 `https://...` 链接。

不要分享本地路径，例如：

- `E:\Codex Project\英语学习\...`
- `file:///E:/Codex%20Project/...`

## 后续更新建议

后续每天生成新页面时：

1. 把当天 HTML 复制一份到 `daily-english/public/`
2. 把 `daily-english/public/index.html` 的跳转日期改成当天
3. `git add .`
4. `git commit -m "Update daily English page for YYYY-MM-DD"`
5. `git push`

GitHub Pages 会自动更新公开网页。
