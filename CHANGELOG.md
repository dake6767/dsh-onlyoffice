# Changelog

## 0.1.7 — 2026-10-21

- **自动保存改为可选（默认关闭）**：`config.autoSave` 默认 `false`，不再在编辑时
  静默写回文件；需要时在插件配置里设 `autoSave: true` 恢复 1.8s 防抖自动保存。
  手动保存不受影响：`⌘S` / `Ctrl+S`、编辑器自带保存按钮、导出下载始终可用。
- `/status` 增加 `autoSave` 与 `saveEnabled` 字段，界面与排查时可直接看到写回策略。

## 0.1.6 — 2026-10-21

- 新增页面内探测函数 `window.__dshooProbe()`：对每个候选 base 分别请求
  `/status` 与 SDK `api.js`，返回 HTTP 状态、content-type、耗时与响应体前缀，
  用于远程部署快速定位“请求打到了哪里”。
- Host base 选择过程写入 console（`[host-base] chosen/candidate failed …`），
  会随下一次诊断上报到 Host 日志。

## 0.1.5 — 2026-10-21

修复：某些 space 页面（`/app/<spaceId>/`）下插件路由不可达。

- 不再假定 Host 路由挂在 origin 根或页面目录：`ensureSdkReady()` 会用
  `GET <base>/status` **依次探测**「页面目录相对」「origin 根」「根相对」三个候选，
  第一个返回合法 JSON 的即被记住，随后 read/save/diagnostic 与 SDK 基址都复用它。
- 这样同一份代码同时适配：桌面壳 `dsh-app://app/`、根路径部署、
  `/app/<spaceId>/` 工作区页面、以及将来真正的反代前缀。
- 版本 0.1.5。

## 0.1.4 — 2026-10-21

- 新增 `DSH_ONLYOFFICE_CACHE_DIR` 环境变量（优先级：插件 config > 环境变量 >
  `$DSH_HOME/onlyoffice-sdk`）。多用户宿主（如 Phalanx 的 test1/test2）可以共用
  一份 SDK 缓存，避免每个用户首次打开都下载约 100 MB。
- SDK 路径中重复的斜杠会被折叠而不是拒绝：演示文稿编辑器会请求
  `sdkjs/slide/themes//themes.js`，之前返回 400 且永不缓存，主题包因此加载失败。
- 新增两个 Host 测试覆盖以上行为。

## 0.1.3 — 2026-10-21

修复：慢速/远程部署下偶发「编辑器渲染超时」，尽管编辑器已正确挂载。

- `hasEditorFrame()` 现在同时识别 **DocsAPI 把编辑器 iframe 挂成兄弟节点**
  （它会替换掉我们传入的 `#containerId` 元素）的情况：先查 id 元素、再查
  `[data-onlyoffice-container-id]` wrapper、最后按 `iframe[name="frameEditor"]`
  的 URL 匹配。之前只在 id 元素内找，慢冷启动时会误判为“没有 iframe”。
- 文档就绪等待从 120s 放宽到 180s（远程首次冷启动可能超过两分钟）。
- `/status` 增加 `pluginVersion`，故障诊断能直接区分「插件版本」和「SDK 版本」；
  环境快照也会说明容器被替换、wrapper 内 iframe 数量。

## 0.1.2 — 2026-10-21

修复：反向代理子路径部署（如 `http://host:18080/app/<hash>/`）下 SDK 资源 404。

- SDK 基址与 Host 路由（status/read/save/diagnostic）改为按 **页面目录**
  （`document.baseURI`）解析，而不是 origin 根路径；这样根路径、
  `/app/<hash>/` 反向代理前缀、以及桌面壳 `dsh-app://app/` 都使用同一套逻辑。
- 新增 `src/client/page-url.js` 纯函数 + 6 个单测覆盖上述三种部署形态
  （含无尾斜杠、query/hash、opaque origin 等边界）。

## 0.1.0 — 2026-10-21

首个公开版本。

### 功能

- 在 DSH 内置文档预览中打开 `.docx` / `.xlsx` / `.pptx`
  （`documentPreviews` 注册 + `sidebar.right.tab.document` 渲染器实现）。
- 可选兼容 dsh-better-sidebar：通过 `ctx.betterSidebar.registerFileViewer`
  注册三个查看器；对服务做加载顺序无关的等待（取 → 可选注入 → 有上限轮询）。
- 编辑并写回原文件：`⌘S` / `Ctrl+S`、编辑器自带保存按钮（`onSave` →
  `ONSAVE`）、以及 1.8 秒防抖的静默自动保存。
- 外部文件变更检测：顶部提示条 + 手动「重新加载」，不会在编辑时悄悄换内容。
- 导出下载、只读/编辑切换、编辑器主题跟随 DSH 明暗。
- 失败自诊断：捕获页面/iframe/console 错误与环境快照，上报
  `POST /dsh-onlyoffice/diagnostic`。

### 引擎与资源

- 内置 [onlyoffice-web-comp](https://github.com/electroluxcode/onlyoffice-web-comp)
  源码（AGPL-3.0，含 9 处补丁，见 `vendor/PATCHES.md`）。
- OnlyOffice 静态 SDK 不随包发布，由 Host 路由按需代理并缓存：
  多镜像 failover（Vercel / jsDelivr / raw.githubusercontent）、磁盘缓存、
  404 负缓存、原子写入。
- 未知 scheme URL 修复（`dsh-app://`）、x2t worker 内联、
  blob worker、保存快捷键/变更事件钩子。

### 安全

- 所有 Host 路由经过 connection fence。
- 写回：工作区包含校验、只允许已存在的 docx/xlsx/pptx、mtime 乐观锁
  （冲突返回 409）、同目录临时文件 + rename 原子替换、保留文件权限。

### 测试

- 87 个 `node:test` 用例：Host 路由/缓存/读写/诊断，Client bundle 结构、
  渲染 smoke、快捷键判定、防抖自动保存、主题 token 白名单、x2t worker 协议。
