# 货位编码工具 · 部署与发版流程

本文说明如何把本工具发布到 **GitHub**，并通过 **GitHub Pages** 生成固定链接；以及**首次发版**与**日后迭代**分别要做什么。适合配合 **GitHub Desktop** 使用（不推荐纯新手先走终端 + Token）。

---

## 一、你会得到什么

| 概念 | 说明 |
|------|------|
| **GitHub 仓库** | 代码存在云端，地址形如 `https://github.com/你的用户名/仓库名` |
| **GitHub Pages** | 同一仓库托管出的网页，地址形如 `https://你的用户名.github.io/仓库名/` |
| **朋友使用方式** | 浏览器打开 Pages 链接即可，**无需登录** |

> 配置数据在各自浏览器的 localStorage，**不会**随你发版自动同步到朋友电脑。

---

## 二、首次发版（只做一次）

### 1. 注册 GitHub

1. 打开 [https://github.com/signup](https://github.com/signup) 注册并验证邮箱。  
2. 记住你的 **用户名**（例如个人主页 `github.com/用户名`）。

### 2. 在网页上创建空仓库

1. 登录 GitHub → 右上角 **`+`** → **New repository**。  
2. **Repository name**：建议英文，如 `huowei-tool`（**不要用单独一个 `-`**，避免链接难记、易和别的教程混淆）。  
3. 选 **Public**。  
4. **不要**勾选 *Add a README*（保持空仓库，减少首次推送冲突）。  
5. 点 **Create repository**。  
6. 记下克隆地址（**Code** → HTTPS），例如：  
   `https://github.com/你的用户名/huowei-tool.git`

### 3. 安装并登录 GitHub Desktop

1. 下载安装：[https://desktop.github.com/](https://desktop.github.com/)  
2. 首次打开 → **Sign in to GitHub.com** → 浏览器登录并 **Authorize**（一般**不需要**在终端里粘贴 Token）。

### 4. 把本地项目加入 Desktop

1. **File** → **Add Local Repository…** → 选择本工具所在文件夹（内含 `index.html`）。  
2. 若提示不是 Git 仓库：按提示 **create a repository** 在该目录初始化，再添加。

### 5. 绑定远程地址（必须和网页仓库一致）

1. **Repository** → **Repository settings…** → **Remote**。  
2. **Primary remote repository** 填步骤 2 里的地址，例如：  
   `https://github.com/你的用户名/huowei-tool.git`  
3. **Save**。

> **常见错误**：远程写成别的仓库名、或占位符 URL，会导致 *repository does not exist*；网页上若没有 `huowei-tool` 只有 `-`，则远程必须是 `https://github.com/用户名/-.git`，二者**名字要一致**。

### 6. 首次推送到 GitHub

1. 若有未提交更改：左下角 **Summary** 写说明 → **Commit to main**。  
2. 点 **Publish branch**（首次）或 **Push origin**（已有远程时）。  
3. 浏览器打开 `https://github.com/你的用户名/huowei-tool`，应能看到 `index.html` 等文件。

### 7. 开启 GitHub Pages

1. 在该仓库页面 → **Settings** → 左侧 **Pages**。  
2. **Source**：**Deploy from a branch**。  
3. **Branch**：**main**，目录 **/ (root)** → **Save**。  
4. 等待 1～3 分钟，页面会显示站点地址，一般为：  
   `https://你的用户名.github.io/huowei-tool/`  
5. 将此链接发给朋友；你本地用 **强制刷新**（Mac：`Cmd + Shift + R`）可看最新静态资源。

**首次发版到此结束。**

---

## 三、未来迭代（每次改完代码）

日常只需在 **GitHub Desktop** 里完成「存档 → 上传」：

1. 用编辑器修改项目文件并**保存**。  
2. 打开 **GitHub Desktop**，选中本仓库。  
3. 在 **Changes** 中查看差异。  
4. 左下角 **Summary** 写本次说明（如：`fix: 教程弹窗`）。  
5. 点 **Commit to main**。  
6. 点 **Push origin**（或顶部同步/推送类按钮）。  
7. 等待 **约 1～2 分钟**（GitHub Pages 构建），朋友**仍打开原链接**，建议 **强制刷新**。

无需重新配置 Pages；除非你在 Settings 里改过分支或关闭了 Pages。

---

## 四、自检清单（出问题先看这里）

| 现象 | 处理方向 |
|------|----------|
| 网页 404 | 仓库名与浏览器地址是否一致；Pages 是否已 Save；是否等了数分钟。 |
| Desktop：仓库不存在 | **Repository settings → Remote** 的 URL 是否与网页仓库 **完全一致**（含仓库名）。 |
| 只有名为 `-` 的仓库 | 远程应用 `.../用户名/-.git`，或新建 `huowei-tool` 并把 Remote 改过去再 Push。 |
| Pages 空白或脚本错 | 浏览器 F12 看控制台；确认 `index.html` 使用相对路径引用 `./styles.css`、`./app.js`。 |
| 推送权限 / 登录 | **GitHub Desktop → Settings → Accounts** 确认已登录正确账号。 |

---

## 五、可选：用终端推送（熟悉 Git 后）

若远程已正确、本机已配置凭据：

```bash
cd "/path/to/货位编码工具"
git add .
git commit -m "说明本次改动"
git push origin main
```

密码处使用 **Personal Access Token**（非网页登录密码）；新手优先用 Desktop，避免凭据与集成终端显示问题。

---

## 六、版本记录建议

**Summary / commit 信息**可简短中文或英文，例如：

- `feat: 新增某某功能`  
- `fix: 修复某某问题`  
- `docs: 更新部署说明`  

便于日后在 **History** 里辨认每次发版做了什么。

---

*文档与仓库内工具同步维护；部署策略以 GitHub 当前页面为准。*
