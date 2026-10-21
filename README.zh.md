# dsh-onlyoffice

[English](./README.md) | **中文**

[![npm version](https://img.shields.io/npm/v/@dake6767/dsh-onlyoffice.svg)](https://www.npmjs.com/package/@dake6767/dsh-onlyoffice)
[![license: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue.svg)](./LICENSE)
[![CI](https://github.com/dake6767/dsh-onlyoffice/actions/workflows/ci.yml/badge.svg)](https://github.com/dake6767/dsh-onlyoffice/actions/workflows/ci.yml)


一个 DeepSeek Harness 插件：在 DSH 侧栏的文件管理里打开 **.docx / .xlsx / .pptx**，用内嵌的 **OnlyOffice** 编辑器查看和编辑。纯前端静态 SDK，**不需要 Document Server，也不需要后端转换服务**。文档始终留在本机，唯一的外网流量是首次打开时下载 OnlyOffice 引擎资源（之后走本机磁盘缓存）。

| | |
|---|---|
| 包名 | `@dake6767/dsh-onlyoffice` |
| Host 半部 | `lib/index.js` — `/dsh-onlyoffice/*` 路由（SDK 代理 + 缓存、读/写、诊断） |
| Client 半部 | `lib/client.js` — 一个 module-loader bundle（约 570 KiB），内含引擎 |
| 编辑器引擎 | [`onlyoffice-web-comp`](https://github.com/electroluxcode/onlyoffice-web-comp) 的源码（放在 `vendor/onlyoffice-web-comp`，有 9 处补丁，见 `vendor/PATCHES.md`） |
| 许可证 | AGPL-3.0（见 [`LICENSE`](./LICENSE)、[`THIRD_PARTY_NOTICES.md`](./THIRD_PARTY_NOTICES.md)） |

## 功能

- **两条集成路径，一个引擎**
  - **DSH 内置文档预览（主路径，无第三方依赖）**：通过 `ctx.documentPreviews` + `sidebar.right.tab.document` 槽注册实现（`loading: 'renderer'`）。
  - **dsh-better-sidebar（可选）**：通过 `ctx.betterSidebar.registerFileViewer` 注册三个文件查看器；对服务做了「先取 → 可选注入 → 有上限轮询」的等待，两个插件无论谁先加载都能组合成功。
- **编辑后写回原文件**
  - 编辑器里按 `⌘S` / `Ctrl+S`（把 vendored 引擎的快捷键拦截器转接到本插件），以及编辑器自带的保存入口（`onSave` → `ONSAVE`）都会写回原文件；
  - **自动保存**：每次文档变更会触发 1.8 秒防抖的静默写回；手动保存会立即冲刷并给出提示。
- **外部修改不会悄悄覆盖你的编辑器**：预览 watcher 发现文件版本变化时，预览顶部出现 *“文件已在磁盘上更新。［重新加载］”*，由你决定何时载入。
- **导出下载**、**只读/编辑切换**、编辑器界面明暗主题跟随 DSH。
- **自诊断**：打开失败时错误界面会列出捕获到的控制台/iframe 错误和环境快照，并上报到 `POST /dsh-onlyoffice/diagnostic`（写入 `<cacheDir>/.requests.log`）。

## 环境要求

- DSH 版本需要提供 `documentPreviews` + `slots` 客户端服务（即自带 `@deepseek-ai/dsh-client-ui-sidebar-documentpreview`）；可选路径需要 `dsh-better-sidebar`。
- 首次打开文档需要能访问下文的 SDK 镜像之一。

## 安装
> 支持反向代理子路径部署：插件所有路由都按页面目录解析，因此 `https://host/app/<id>/` 与 `http://127.0.0.1:53803/`、桌面壳 `dsh-app://app/` 表现一致。


### 方式一：本地目录链接（开发/自用，当前方式）

```bash
# 在 DSH 里用 plugin_manager:
#   action: install_bundle
#   target: /Users/wuliqi/projects/dsh-onlyoffice
```

会向 profile 添加依赖并启用 bundle 行。**客户端改动刷新页面（⌘R）即可；Host 半部改动需要完整重启 DSH**（目录链接安装下 Host 模块的 HMR 不可靠）。

### 方式二：npm 包

```bash
npm i @dake6767/dsh-onlyoffice
# 或用 DSH 的插件管理器安装该包
```

## 构建与测试

```bash
node build/build-client.mjs   # vendor + src/client -> lib/client.js（约 570 KiB）
node build/build-host.mjs     # host/*.js -> lib/*.js
node --test "test/**/*.test.mjs"   # 87 个测试
```

## OnlyOffice 静态 SDK 从哪来

插件**不打包 SDK**（完整目录约 1.96 万个文件 / 1.1 GB）。Host 按需代理单个文件并缓存：

```
GET /dsh-onlyoffice/sdk/onlyoffice/<version>/<rest>
  命中缓存 -> <cacheDir>/<version>/<rest>
  未命中   -> <mirror>/onlyoffice/<version>/<rest>（按顺序取第一个可用的镜像）
```

默认镜像（网络错误 / 超时 / 404 / 5xx 时自动切换）：

1. `https://onlyoffice-web-comp.vercel.app/packages`
2. `https://cdn.jsdelivr.net/gh/electroluxcode/onlyoffice-web-comp@main/public/packages`
3. `https://raw.githubusercontent.com/electroluxcode/onlyoffice-web-comp/main/public/packages`

明确 404 的资源（例如所有镜像都没有的 `dictionaries/en_US/*` 拼写词典）会被记住 10 分钟，避免每次渲染都重走一遍镜像链。

### 配置（`cordis.patch.yml` 的 `onlyoffice` 行）

```yaml
- id: onlyoffice
  name: '@dake6767/dsh-onlyoffice'
  config:
    version: '9.4.0-develop'          # SDK 目录版本（客户端从 /status 读取，不写死）
    upstreamBase: 'https://onlyoffice-web-comp.vercel.app/packages'
    upstreamFallbacks: []             # [] = 只用单一源；省略则用上面的默认镜像链
    cacheDir: '~/.dsh/onlyoffice-sdk'
    fetchTimeoutMs: 300000            # 单个资源的上游超时
    saveEnabled: true                 # false 关闭写回路由
    maxSaveBytes: 268435456           # 写回大小上限（256 MiB）
```

**离线/内网**：把同样目录结构的文件预先放进 `cacheDir`，或把 `upstreamBase` 指向自建静态站点。

## HTTP 路由（全部经过 connection fence，与 `/sidebar/*` 一致）

| 路由 | 作用 |
|---|---|
| `GET /status` | `{version, sdkBaseUrl, upstreams, cacheDir, files, bytes, lastError}` |
| `GET /sdk/onlyoffice/<ver>/<rest>` | SDK 代理 + 磁盘缓存（immutable 响应头） |
| `GET /read?sessionId&path[&cwd]` | 读取文档字节，返回 `x-dsh-file-version`（renderer 模式） |
| `POST /save?sessionId&path[&cwd][&mtimeMs]` | 原始文档字节原子写回 |
| `POST /diagnostic` | 浏览器诊断（限长），追加到 `.requests.log` |

**写回安全**（本插件唯一的写路径）：

- 目标必须**已存在**且是常规文件，并且位于会话工作区内（优先 `fs.contains`，回退 realpath 包含判断）；
- 扩展名只允许 `docx / xlsx / pptx`（引擎原生输出格式；把 docx 字节写回 `.odt/.doc/.csv` 会损坏文件，这些格式只读 + 导出）；
- 可选 `mtimeMs` 作为乐观锁：磁盘文件已变化时返回 **409 conflict**，而不是覆盖；
- 同目录临时文件 + `rename` 原子替换，保留原文件权限位。

## 使用说明

1. 右侧栏「文件」→ 点开 docx/xlsx/pptx → 自动用 OnlyOffice 打开（预览右上角可看到「OnlyOffice 文档」实现）；
2. 编辑后 **`⌘S` / `Ctrl+S`** 保存；不按也行，停手约 2 秒会自动保存（静默）；
3. 点 OnlyOffice 自带软盘同样会写回；
4. 文件被外部修改（比如 Agent 写的）→ 顶部提示条 →「重新加载」查看最新内容；
5. 「导出下载」把当前文档导出为本地文件；「编辑/只读」切换编辑权限。

> 注意：编辑器**内部**的“所有更改已保存”只代表它自己的内存服务状态；真正的落盘由本插件负责（⌘S、软盘、自动保存三条路径都走同一个写回逻辑）。

## 已知限制

- **没有实时协同/字符级热更新**：静态 SDK 没有协同通道，真正的共同编辑需要 Document Server / WOPI 后端。外部修改是“提示 → 你点击 → 整篇重载”；
- 写回只支持原生 OOXML（docx/xlsx/pptx），其余格式只读 + 导出；
- `dsh-resource://file/absolute/...` 这类地址没有会话信息，可查看但无法写回；
- x2t 转换是页面级单例，已用锁串行化；首次打开需要下载对应 SDK 片段；
- 侧栏的刷新控件属于内置预览；本插件额外提供的唯一重载入口是“外部更新”提示条。

## 排错

| 现象 | 排查 |
|---|---|
| 首次打开很久 / 一直显示「加载引擎」 | 冷启动在下载 SDK；观察 `<cacheDir>` 的体积是否增长 |
| 「编辑器渲染超时」 | 错误详情里有控制台错误 + iframe 状态快照 |
| `dictionaries/en_US/*` 404 | 无语义影响（上游没有拼写词典），已做负缓存 |
| 保存报 `conflict` | 文件在磁盘上变了——先重新加载再保存 |
| 保存报 `method not allowed` | 运行中的 Host 还是旧代码——完整重启 DSH |
| 插件更新后侧栏空白 | 客户端 bundle 崩溃；错误界面会显示 `[console-*]`，同时看 `.requests.log` |

`<cacheDir>/.requests.log` 记录每次 SDK/读/写/诊断请求的状态码与耗时（2 MiB 轮转），是定位“浏览器到底发了什么请求”的最快入口。

## 目录结构

```
host/                 Host 半部（只依赖 node:* 内置模块）
  index.js            路由注册、SDK 代理、状态、诊断、日志
  save.js             /read 与 /save：工作区校验、原子写、乐观锁
  sdk-cache.js        纯函数助手（路径映射、内容类型、镜像链、缓存统计）
src/client/           Client 半部（esbuild 打成单文件 bundle）
  index.jsx           注册两条集成路径；装载样式/诊断/快捷键
  native.jsx          内置预览的 documentPreviews 实现 + resource 地址解析
  office-view.jsx     OnlyOffice 视图：加载/编辑/保存/自动保存/外部更新提示
  sdk.js              /status、/read、/save 客户端调用；SDK 资源基址注册
  autosave.js         防抖自动保存
  diagnostics.js      页面/iframe/console 捕获 + 环境快照
build/                esbuild 构建（处理 x2t worker 内联）
vendor/onlyoffice-web-comp/   vendored 引擎（补丁见 vendor/PATCHES.md）
test/                 node:test：Host 路由/缓存/读写 + Client bundle/渲染 smoke
fixtures/             演示文档（docx/xlsx/pptx + live-sync 演示）
```

## 许可证与第三方

- 本包：**AGPL-3.0**（[`LICENSE`](./LICENSE)），源码中带 `SPDX-License-Identifier: AGPL-3.0-only`；
- vendored 引擎：上游同样 AGPL-3.0，其许可文本保留在 `vendor/onlyoffice-web-comp/LICENSE`，修改记录见 [`vendor/PATCHES.md`](./vendor/PATCHES.md)；
- OnlyOffice SDK 资产**不在包内**，运行时从镜像下载；其许可情况（社区版 AGPL、Developer Edition 商业版、`x2t` WASM 许可不明确）见 [`THIRD_PARTY_NOTICES.md`](./THIRD_PARTY_NOTICES.md)；
- 内置 brotli 解码器保留 Apache-2.0 头；
- OnlyOffice 是 Ascensio System SIA 的商标，本项目是独立集成，与其无隶属或背书关系。
