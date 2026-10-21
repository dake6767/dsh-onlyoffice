# Changelog

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
