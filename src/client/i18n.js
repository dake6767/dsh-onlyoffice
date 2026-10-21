// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * In-UI copy for the dsh-onlyoffice viewers.
 *
 * The host's locale service (`ctx.locale`) is in the plugin's `inject` list,
 * and every read here stays defensive so a host without it still activates the
 * viewers. `apply()` subscribes to the service when
 * it exists and re-publishes the active language to `t()`.
 */

const DICT = {
  zh: {
    viewerOnlyOffice: 'OnlyOffice 文档',
    viewerDocx: 'Word 文档',
    viewerXlsx: 'Excel 表格',
    viewerPptx: 'PowerPoint 演示文稿',

    loadingEngine: '正在加载 OnlyOffice 引擎…',
    openingDocument: '正在打开文档…',
    converting: '正在转换文档…',
    slowHint: '首次打开需要下载并初始化 OnlyOffice 引擎，可能需要一点时间。',

    retry: '重试',
    reload: '重新加载',
    save: '保存',
    saving: '保存中…',
    savedNotice: '已保存到原文件。',
    saveFailed: '保存失败',
    saveConflict: '文件已在磁盘上被修改，请先「重新加载」再保存',
    saveHint: '把编辑后的文档写回原文件',
    saveReadOnly: '当前是只读模式，切换到编辑后再保存',
    externalChanged: '文件已在磁盘上更新。',
    externalReload: '重新加载',
    reloadFailed: '重新加载失败',
    saveNoTarget: '没有可写回的原始文件路径（仅支持从文件树打开的文档）',
    saveUnsupported: '该格式不支持直接写回（仅 docx / xlsx / pptx），请用「导出下载」',
    export: '导出下载',
    exporting: '导出中…',
    readOnly: '只读',
    editable: '编辑',
    toggleReadOnly: '切换只读 / 编辑',
    openExternal: '在浏览器中打开',
    dismiss: '关闭',
    exportingHint: '正在导出，请稍候…',

    loadFailed: '无法打开该文档',
    errorStatus: '无法连接 Office 预览服务（/dsh-onlyoffice/status）。',
    errorStatusHint: '宿主端 dsh-onlyoffice 插件可能未启用，或该路由暂时不可用。',
    errorNoMedia: '没有可用的文件地址（mediaUrl 缺失）。',
    errorUnsupported: '当前环境不支持 File API，无法打开 Office 文档。',
    errorFetch: '无法读取文件内容。',
    errorEngine: 'OnlyOffice 引擎初始化失败。',
    errorConvert: '文档转换失败。',
    errorReady: '编辑器渲染超时。',
    errorTimeout: '加载超时。',
    errorTimeoutHint: '宿主端 SDK 路由可能不可达；可重试，或确认 dsh-onlyoffice 插件已启用。',
    errorRetryHint: '可以重试；若持续失败，请在文件树中下载该文件后用本地 Office 打开。',
    exportFailed: '导出失败',
    readonlyFailed: '切换只读状态失败',
  },
  en: {
    viewerOnlyOffice: 'OnlyOffice document',
    viewerDocx: 'Word document',
    viewerXlsx: 'Excel spreadsheet',
    viewerPptx: 'PowerPoint presentation',

    loadingEngine: 'Loading the OnlyOffice engine…',
    openingDocument: 'Opening the document…',
    converting: 'Converting the document…',
    slowHint: 'The first open downloads and initializes the OnlyOffice engine, so it may take a while.',

    retry: 'Retry',
    reload: 'Reload',
    save: 'Save',
    saving: 'Saving…',
    savedNotice: 'Saved back to the original file.',
    saveFailed: 'Save failed',
    saveConflict: 'The file changed on disk — reload the document before saving',
    saveHint: 'Write the edited document back to its file',
    saveReadOnly: 'The document is read-only; switch to editing before saving',
    externalChanged: 'The file changed on disk.',
    externalReload: 'Reload',
    reloadFailed: 'Reload failed',
    saveNoTarget: 'No original file path to save to (open the document from the file tree)',
    saveUnsupported: 'This format cannot be written back (docx / xlsx / pptx only); use Export instead',
    export: 'Export',
    exporting: 'Exporting…',
    readOnly: 'Read-only',
    editable: 'Editing',
    toggleReadOnly: 'Toggle read-only / editing',
    openExternal: 'Open in browser',
    dismiss: 'Dismiss',
    exportingHint: 'Exporting, please wait…',

    loadFailed: 'Could not open this document',
    errorStatus: 'Cannot reach the Office preview service (/dsh-onlyoffice/status).',
    errorStatusHint: 'The host-side dsh-onlyoffice plugin may be disabled, or the route is temporarily unavailable.',
    errorNoMedia: 'No file URL is available (mediaUrl is missing).',
    errorUnsupported: 'This environment has no File API, so Office documents cannot be opened.',
    errorFetch: 'Could not read the file contents.',
    errorEngine: 'The OnlyOffice engine failed to initialize.',
    errorConvert: 'The document conversion failed.',
    errorReady: 'The editor did not finish rendering in time.',
    errorTimeout: 'Loading timed out.',
    errorTimeoutHint: 'The host SDK route may be unreachable; retry, or check that the dsh-onlyoffice plugin is enabled.',
    errorRetryHint: 'You can retry; if it keeps failing, download the file from the tree and open it locally.',
    exportFailed: 'Export failed',
    readonlyFailed: 'Could not switch the read-only state',
  },
};

const FALLBACK_LANG = 'en';

let activeLang = detectLang();
const listeners = new Set();

function detectLang() {
  try {
    const nav = typeof navigator === 'undefined' ? undefined : navigator;
    const raw = String((nav && (nav.language || nav.languages?.[0])) || '');
    if (/^zh\b/i.test(raw)) return 'zh';
    if (raw) return 'en';
  } catch {
    /* ignore */
  }
  return FALLBACK_LANG;
}

/** Read the active locale id from a host locale service (defensive). */
function readLocaleLang(locale) {
  try {
    const snapshot = locale && typeof locale.getSnapshot === 'function' ? locale.getSnapshot() : undefined;
    const id = snapshot && typeof snapshot.active === 'string' ? snapshot.active : '';
    if (/^zh/i.test(id)) return 'zh';
    if (/^en/i.test(id)) return 'en';
    if (/^ja/i.test(id)) return 'zh';
  } catch {
    /* ignore */
  }
  return undefined;
}

function lookup(lang, key) {
  const table = DICT[lang] || DICT[FALLBACK_LANG];
  return table[key] ?? DICT[FALLBACK_LANG][key] ?? key;
}

/** Translate one key for the active language; `{name}` placeholders are filled from params. */
export function t(key, params) {
  let value = lookup(activeLang, key);
  if (params) {
    for (const name of Object.keys(params)) {
      value = value.split(`{${name}}`).join(String(params[name]));
    }
  }
  return value;
}

export function getLang() {
  return activeLang;
}

export function subscribeLang(listener) {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function setLang(next) {
  if (!next || next === activeLang) return;
  activeLang = next;
  for (const listener of [...listeners]) {
    try {
      listener(activeLang);
    } catch {
      /* a subscriber must not break the others */
    }
  }
}

/**
 * Attach the host locale service when present: track its active language and
 * surface it through `t()`. Returns a disposer.
 */
export function attachLocale(ctx) {
  const locale = ctx && ctx.locale;
  if (!locale || typeof locale.subscribe !== 'function') return () => {};

  const sync = () => {
    const lang = readLocaleLang(locale);
    if (lang) setLang(lang);
  };

  sync();

  let unsubscribe;
  try {
    unsubscribe = locale.subscribe(sync);
  } catch {
    return () => {};
  }

  return () => {
    try {
      unsubscribe?.();
    } catch {
      /* ignore */
    }
  };
}
