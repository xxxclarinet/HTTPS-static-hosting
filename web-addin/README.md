# 班次邮件分发助手（Outlook Web 加载项 · 方案甲）

一个**纯网页**的 Outlook 加载项：不依赖 EXE、不依赖 .NET、不需要 Azure 应用注册或管理员授权。
它在 Outlook（网页版 / 新版 Outlook for Windows / 经典 Outlook）里以任务窗格运行，帮你：

1. 在**源班次邮件**上解析每位组员与对应班次图，生成分发清单；
2. 在**新建邮件**里逐封自动填好收件人 / 主题 / 正文，并**内嵌该组员的班次图**；
3. 你按 `Ctrl+S` 存为草稿，统一复核后再手动发送 —— 全程不自动发送，无误发风险。

> 数据安全：所有邮件/图片数据只在你 Outlook 的网页里处理，**不外传**。托管在 GitHub Pages 上的只是**不含任何数据的程序代码**。

---

## 一、部署（托管一次，全员共用）

托管地址固定为：`https://xxxclarinet.github.io/HTTPS-static-hosting/`
（已写进 `manifest.xml`。若你换了仓库/用户名，需同步替换 manifest 与本文件里的这个地址。）

1. 把 **`web-addin/` 目录里的全部内容**（保持子目录结构）推送到仓库 `HTTPS-static-hosting` 的**根目录**：
   ```
   manifest.xml
   taskpane.html  taskpane.css  taskpane.js  config.js
   domain/  storage/  outlook/  assets/
   ```
2. 在仓库 **Settings → Pages**，Source 选 `Deploy from a branch`，分支选 `main` / 目录 `/ (root)`，保存。
3. 等一两分钟，访问 `https://xxxclarinet.github.io/HTTPS-static-hosting/taskpane.html`，
   能看到加载项页面（而不是 404 / 下载）即部署成功。

> 图标：`assets/` 下已含 `icon-16/32/80.png`（占位）。想换自己的图标，替换同名文件即可，
> 或重跑 `assets/make-icons.ps1` 重新生成。

---

## 二、安装到 Outlook（每位同事一次性 sideload）

**网页版 / 新版 Outlook for Windows：**

1. 设置（齿轮）→ **常规 → 管理加载项**（或 “Get Add-ins”）。
2. 选 **My add-ins → Add a custom add-in → Add from URL**。
3. 填：`https://xxxclarinet.github.io/HTTPS-static-hosting/manifest.xml` → 安装。

> 也可下载 `manifest.xml` 后用 “Add from file” 安装。把这个小文件放公司网盘发给同事即可。
> 若贵司关闭了自定义加载项 sideload，则需 IT 在管理中心集中部署该 manifest。

---

## 三、使用流程

**第 1 步（阅读模式）：生成分发清单**

1. 打开那封**源班次邮件**。
2. 工具栏点 **“班次分发助手”** 打开任务窗格。
3. （首次）展开底部“配置”，设置邮件域名、主题模板、正文、落款，保存。
4. 点 **“解析源邮件”**：窗格列出每位组员（Login 可现场修改）、收件人、主题与班次图缩略图，并提示无法识别的项。
5. 确认无误后点 **“保存为分发清单”**。

**第 2 步（撰写模式）：逐封生成草稿**

1. **新建一封空白邮件**。
2. 在撰写窗口工具栏点 **“班次分发助手”**（建议把窗格**固定/钉住**，省得每封都点开）。
3. 点 **“填充当前邮件（下一位）”**：自动填好收件人/主题/正文并内嵌班次图。
4. 核对后按 **`Ctrl+S`** 存草稿，关闭；**再新建下一封**，重复，直到窗格显示“待处理 0 位”。
5. 最后到**草稿箱**统一复核，逐封发送。

---

## 四、本版范围与限制（MVP）

已实现：
- 周期推断（含跨年）、Login 合法性校验、主题/收件人/正文渲染、独立图模式按文件名配人、内嵌班次图、配置持久化（roamingSettings）、清单持久化（IndexedDB）。

**暂未实现**（按需后续迭代）：
- 静默批量建草稿（纯 Office.js 做不到，故采用“逐封填充 + 手动 Ctrl+S”）；
- 模板切分模式（整张大图按网格裁剪）、OCR 配人、相邻文本(caption)配人（当前以**文件名**为主要依据）；
- 历史记录 / 同周期去重、定时与新邮件自动触发、失败重试统计、CSV 导出。

已知注意点：
- 内嵌图依赖撰写模式的 `addFileAttachmentFromBase64Async`（Mailbox 需求集 1.8，新版/网页 Outlook 支持）。
- 阅读模式存的清单与撰写模式读取，靠同源 IndexedDB 共享。**若在新版 Outlook 撰写窗口里看不到清单**，多半是撰写窗口与阅读窗格不同源/不同分区所致——请改用**网页版 OWA** 测试该链路，或反馈我再调整传递方式。
- 班次图按文件名（如 `ytshunyu.png`）识别 Login。文件名含多个像 Login 的词会被标为“需人工确认”，可在清单里直接改 Login。
