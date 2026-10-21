window.__ModuleLoader__.load({ id: "@dake6767/dsh-onlyoffice", factory: (require) => {
var module = { exports: {} }; var exports = module.exports;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/client/index.jsx
var index_exports = {};
__export(index_exports, {
  apply: () => apply,
  inject: () => inject,
  viewerDescriptors: () => viewerDescriptors
});
module.exports = __toCommonJS(index_exports);
var import_react3 = require("react");

// src/client/office-view.jsx
var import_react = require("react");

// vendor/onlyoffice-web-comp/internal/editor/types.ts
var X2T_CSV_ENCODING_UTF8 = 46;
var X2T_CSV_ENCODING_GBK = 18;
var X2T_CSV_DELIMITER_TAB = 1;
var X2T_CSV_DELIMITER_SEMICOLON = 2;
var X2T_CSV_DELIMITER_COMMA = 4;
var OFFICE_XML_SIZE_LIMIT_ERROR_MESSAGE = "文件过大，不支持解析";
var OfficeXmlSizeLimitExceededError = class extends Error {
  payload;
  constructor(payload) {
    super(OFFICE_XML_SIZE_LIMIT_ERROR_MESSAGE);
    this.name = "OfficeXmlSizeLimitExceededError";
    this.payload = payload;
  }
};
function isOfficeXmlSizeLimitExceededError(error) {
  return error instanceof OfficeXmlSizeLimitExceededError || error instanceof Error && error.name === "OfficeXmlSizeLimitExceededError" && "payload" in error;
}

// vendor/onlyoffice-web-comp/const/index.ts
var ONLYOFFICE_ID = "iframe-office-id";
var ONLYOFFICE_CONTAINER_CONFIG = {
  PARENT_SELECTOR: ".onlyoffice-container",
  PARENT_CLASS_NAME: "onlyoffice-container",
  STYLE: {
    position: "absolute",
    inset: 0
  }
};
var ONLYOFFICE_EVENT_KEYS = {
  SAVE_DOCUMENT: "saveDocument",
  DOCUMENT_READY: "documentReady",
  LOADING_CHANGE: "loadingChange",
  ONSAVE: "onSave",
  OFFICE_XML_SIZE_LIMIT_EXCEEDED: "officeXmlSizeLimitExceeded"
};
var ONLYOFFICE_LANG_KEY = {
  ZH: "zh",
  EN: "en"
};
var OFFICE_THEME = {
  LIGHT: "theme-light",
  CLASSIC_LIGHT: "theme-classic-light",
  WHITE: "theme-white",
  DARK: "theme-dark",
  NIGHT: "theme-night",
  CONTRAST_DARK: "theme-contrast-dark"
};
var DEFAULT_OFFICE_THEME = OFFICE_THEME.WHITE;
var OFFICE_THEME_OPTIONS = [
  { id: OFFICE_THEME.WHITE, label: "浅色" },
  { id: OFFICE_THEME.CLASSIC_LIGHT, label: "经典浅色" },
  { id: OFFICE_THEME.LIGHT, label: "Light" },
  { id: OFFICE_THEME.DARK, label: "深色" },
  { id: OFFICE_THEME.NIGHT, label: "夜间" },
  { id: OFFICE_THEME.CONTRAST_DARK, label: "高对比深色" }
];
var OFFICE_XML_EVENT_CONFIG = {
  default: {
    isEnable: false,
    limitBytes: 1024 * 1024 * 2048
  }
};
var ASC_RESTRICTION_NONE = 0;
var ASC_RESTRICTION_VIEW = 128;
var X2T_PDF_DEFAULT_FONT_FILE = "Carlito-Regular.ttf";
var DEFAULT_ONLYOFFICE_VERSION = "9.4.0-develop";
var DEFAULT_ONLYOFFICE_ROOT = `/packages/onlyoffice/${DEFAULT_ONLYOFFICE_VERSION}`;
var staticResourceOptions = null;
function trimTrailingSlash(value) {
  return value.replace(/\/+$/, "");
}
function buildStaticResource() {
  const apiJs = "/web-apps/apps/api/documents/api.js";
  const preloadHtml = "/web-apps/apps/api/documents/preload.html";
  const cdnOrigin = staticResourceOptions?.cdnOrigin ? trimTrailingSlash(staticResourceOptions.cdnOrigin) : "";
  const onlyofficeVersion = staticResourceOptions?.onlyofficeVersion || DEFAULT_ONLYOFFICE_VERSION;
  const onlyofficeRoot = cdnOrigin ? `${cdnOrigin}/onlyoffice/${onlyofficeVersion}` : DEFAULT_ONLYOFFICE_ROOT;
  const x2tRoot = `${onlyofficeRoot}/x2t`;
  const x2tPdfFontsRoot = `${onlyofficeRoot}/x2t-fonts`;
  return {
    version: {
      onlyofficeSdk: onlyofficeRoot,
      x2t: x2tRoot
    },
    onlyoffice: {
      root: onlyofficeRoot,
      apiJs,
      preloadHtml,
      apiUrl: onlyofficeRoot + apiJs,
      preloadUrl: onlyofficeRoot + preloadHtml
    },
    x2t: {
      root: x2tRoot,
      script: `${x2tRoot}/x2t.js`,
      wasm: `${x2tRoot}/x2t.wasm`,
      pdfFonts: {
        root: x2tPdfFontsRoot,
        default: `${x2tPdfFontsRoot}/${X2T_PDF_DEFAULT_FONT_FILE}`
      }
    }
  };
}
var staticResourceCache = null;
function registerOnlyOfficeStaticResource(options) {
  staticResourceOptions = { ...options };
  staticResourceCache = null;
  return getStaticResource();
}
function resetOnlyOfficeStaticResource() {
  staticResourceOptions = null;
  staticResourceCache = null;
  return getStaticResource();
}
function getStaticResource() {
  if (!staticResourceCache) {
    staticResourceCache = buildStaticResource();
  }
  return staticResourceCache;
}
function isOnlyOfficeCdnMode() {
  const root = getStaticResource().onlyoffice.root;
  if (!/^https?:\/\//i.test(root) || typeof window === "undefined") {
    return false;
  }
  try {
    return new URL(root).origin !== window.location.origin;
  } catch {
    return false;
  }
}
var STATIC_RESOURCE = {
  get version() {
    return getStaticResource().version;
  },
  get onlyoffice() {
    return getStaticResource().onlyoffice;
  },
  get x2t() {
    return getStaticResource().x2t;
  }
};
function resolveSiteUrl(origin, path) {
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(path)) {
    return path;
  }
  return `${origin}${path.startsWith("/") ? path : `/${path}`}`;
}
var FILE_TYPE = {
  DOCX: "DOCX",
  XLSX: "XLSX",
  PPTX: "PPTX"
};
var AppType = {
  word: 1,
  slide: 3,
  cell: 2,
  draw: 4,
  pdf: 5
};
var docTypeMap = {
  docx: AppType.word,
  doc: AppType.word,
  odt: AppType.word,
  rtf: AppType.word,
  txt: AppType.word,
  html: AppType.word,
  mht: AppType.word,
  epub: AppType.word,
  fb2: AppType.word,
  mobi: AppType.word,
  docm: AppType.word,
  dotx: AppType.word,
  dotm: AppType.word,
  oform: AppType.word,
  docxf: AppType.word,
  pptx: AppType.slide,
  ppt: AppType.slide,
  odp: AppType.slide,
  ppsx: AppType.slide,
  pptm: AppType.slide,
  ppsm: AppType.slide,
  potx: AppType.slide,
  potm: AppType.slide,
  otp: AppType.slide,
  odg: AppType.slide,
  xlsx: AppType.cell,
  xls: AppType.cell,
  ods: AppType.cell,
  csv: AppType.cell,
  xlsm: AppType.cell,
  xltx: AppType.cell,
  xltm: AppType.cell,
  xlsb: AppType.cell,
  ots: AppType.cell,
  vsdx: AppType.draw,
  vssx: AppType.draw,
  vstx: AppType.draw,
  vsdm: AppType.draw,
  vssm: AppType.draw,
  vstm: AppType.draw,
  pdf: AppType.pdf
};
var appTypeName = {
  [AppType.word]: "word" /* Word */,
  [AppType.slide]: "slide" /* Slide */,
  [AppType.cell]: "cell" /* Cell */,
  [AppType.draw]: "draw" /* Draw */,
  [AppType.pdf]: "pdf" /* Pdf */
};
function getDocumentType(ext) {
  const code = docTypeMap[ext.toLowerCase()];
  if (code === void 0) {
    return "word" /* Word */;
  }
  return appTypeName[code] ?? "word" /* Word */;
}

// vendor/onlyoffice-web-comp/internal/editor/utils.ts
function getFileExt(name) {
  const type = name.split(".").pop() || "";
  return type.toLowerCase();
}
var x2tSourceFormatByExt = {
  docx: 65 /* AVS_FILE_DOCUMENT_DOCX */,
  doc: 66 /* AVS_FILE_DOCUMENT_DOC */,
  odt: 67 /* AVS_FILE_DOCUMENT_ODT */,
  rtf: 68 /* AVS_FILE_DOCUMENT_RTF */,
  txt: 69 /* AVS_FILE_DOCUMENT_TXT */,
  html: 70 /* AVS_FILE_DOCUMENT_HTML */,
  mht: 71 /* AVS_FILE_DOCUMENT_MHT */,
  epub: 72 /* AVS_FILE_DOCUMENT_EPUB */,
  fb2: 73 /* AVS_FILE_DOCUMENT_FB2 */,
  mobi: 74 /* AVS_FILE_DOCUMENT_MOBI */,
  docm: 75 /* AVS_FILE_DOCUMENT_DOCM */,
  dotx: 76 /* AVS_FILE_DOCUMENT_DOTX */,
  dotm: 77 /* AVS_FILE_DOCUMENT_DOTM */,
  fodt: 78 /* AVS_FILE_DOCUMENT_ODT_FLAT */,
  ott: 79 /* AVS_FILE_DOCUMENT_OTT */,
  oform: 85 /* AVS_FILE_DOCUMENT_OFORM */,
  docxf: 86 /* AVS_FILE_DOCUMENT_DOCXF */,
  xlsx: 257 /* AVS_FILE_SPREADSHEET_XLSX */,
  xls: 258 /* AVS_FILE_SPREADSHEET_XLS */,
  ods: 259 /* AVS_FILE_SPREADSHEET_ODS */,
  csv: 260 /* AVS_FILE_SPREADSHEET_CSV */,
  xlsm: 261 /* AVS_FILE_SPREADSHEET_XLSM */,
  xltx: 262 /* AVS_FILE_SPREADSHEET_XLTX */,
  xltm: 263 /* AVS_FILE_SPREADSHEET_XLTM */,
  xlsb: 264 /* AVS_FILE_SPREADSHEET_XLSB */,
  fods: 265 /* AVS_FILE_SPREADSHEET_ODS_FLAT */,
  ots: 266 /* AVS_FILE_SPREADSHEET_OTS */,
  pptx: 129 /* AVS_FILE_PRESENTATION_PPTX */,
  ppt: 130 /* AVS_FILE_PRESENTATION_PPT */,
  odp: 131 /* AVS_FILE_PRESENTATION_ODP */,
  ppsx: 132 /* AVS_FILE_PRESENTATION_PPSX */,
  pptm: 133 /* AVS_FILE_PRESENTATION_PPTM */,
  ppsm: 134 /* AVS_FILE_PRESENTATION_PPSM */,
  potx: 135 /* AVS_FILE_PRESENTATION_POTX */,
  potm: 136 /* AVS_FILE_PRESENTATION_POTM */,
  fodp: 137 /* AVS_FILE_PRESENTATION_ODP_FLAT */,
  otp: 138 /* AVS_FILE_PRESENTATION_OTP */,
  pdf: 513 /* AVS_FILE_CROSSPLATFORM_PDF */,
  pdfa: 521 /* AVS_FILE_CROSSPLATFORM_PDFA */,
  djvu: 515 /* AVS_FILE_CROSSPLATFORM_DJVU */,
  xps: 516 /* AVS_FILE_CROSSPLATFORM_XPS */,
  vsdx: 16385 /* AVS_FILE_DRAW_VSDX */,
  vssx: 16386 /* AVS_FILE_DRAW_VSSX */,
  vstx: 16387 /* AVS_FILE_DRAW_VSTX */,
  vsdm: 16388 /* AVS_FILE_DRAW_VSDM */,
  vssm: 16389 /* AVS_FILE_DRAW_VSSM */,
  vstm: 16390 /* AVS_FILE_DRAW_VSTM */,
  zip: 2048 /* AVS_FILE_OTHER */ + 9,
  json: 2056 /* AVS_FILE_OTHER_JSON */
};
function getDefaultX2tSourceFormat(fileType) {
  switch (getDocumentType(fileType)) {
    case "cell" /* Cell */:
      return 257 /* AVS_FILE_SPREADSHEET_XLSX */;
    case "slide" /* Slide */:
      return 129 /* AVS_FILE_PRESENTATION_PPTX */;
    case "draw" /* Draw */:
      return 16385 /* AVS_FILE_DRAW_VSDX */;
    default:
      return 65 /* AVS_FILE_DOCUMENT_DOCX */;
  }
}
function getX2tBinFormat(fileType) {
  switch (getDocumentType(fileType)) {
    case "cell" /* Cell */:
      return 8194 /* AVS_FILE_CANVAS_SPREADSHEET */;
    case "slide" /* Slide */:
      return 8195 /* AVS_FILE_CANVAS_PRESENTATION */;
    case "draw" /* Draw */:
      return 8192 /* AVS_FILE_CANVAS */ + 5;
    default:
      return 8193 /* AVS_FILE_CANVAS_WORD */;
  }
}
function getX2tExportBinFormat(fileType) {
  switch (getDocumentType(fileType)) {
    case "slide" /* Slide */:
      return 4099 /* AVS_FILE_TEAMLAB_PPTY */;
    default:
      return getX2tBinFormat(fileType);
  }
}
function getX2tConvertFormats(fileType) {
  const ext = getFileExt(fileType);
  const formatFrom = x2tSourceFormatByExt[ext] ?? getDefaultX2tSourceFormat(fileType);
  return {
    formatFrom,
    formatTo: getX2tBinFormat(fileType)
  };
}
function isValidUtf8(bytes) {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}
function decodeCsvBytes(bytes, encoding) {
  const withoutBom = encoding === X2T_CSV_ENCODING_UTF8 && bytes.length >= 3 && bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191 ? bytes.subarray(3) : bytes;
  if (encoding === X2T_CSV_ENCODING_UTF8) {
    return new TextDecoder("utf-8").decode(withoutBom);
  }
  try {
    return new TextDecoder("gbk").decode(withoutBom);
  } catch {
    return new TextDecoder("latin1").decode(withoutBom);
  }
}
function decodeCsvSample(buffer, encoding) {
  const bytes = new Uint8Array(buffer);
  const sample = bytes.subarray(0, Math.min(bytes.length, 8192));
  return decodeCsvBytes(sample, encoding);
}
function decodeCsvBuffer(buffer, encoding) {
  return decodeCsvBytes(new Uint8Array(buffer), encoding);
}
function encodeCsvBuffer(text, withUtf8Bom) {
  const encoded = new TextEncoder().encode(text);
  if (!withUtf8Bom) {
    return encoded.slice().buffer;
  }
  const withBom = new Uint8Array(encoded.length + 3);
  withBom.set([239, 187, 191], 0);
  withBom.set(encoded, 3);
  return withBom.slice().buffer;
}
function parseCsvText(text, delimiter) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = "";
      continue;
    }
    if (ch === "\r") {
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
      if (text[i + 1] === "\n") {
        i++;
      }
      continue;
    }
    if (ch === "\n") {
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
      continue;
    }
    field += ch;
  }
  row.push(field);
  if (row.length > 1 || row[0] !== "") {
    rows.push(row);
  }
  return rows;
}
function serializeCsvRow(fields, delimiter) {
  return fields.map((field) => serializeCsvField(field, delimiter)).join(delimiter);
}
function serializeCsvField(field, delimiter) {
  if (field.includes(delimiter) || field.includes('"') || field.includes("\n") || field.includes("\r")) {
    return `"${field.replace(/"/g, '""')}"`;
  }
  return field;
}
function toCsvFormulaCell(value) {
  if (/^=".*"$/.test(value)) {
    return value;
  }
  return `="${value.replace(/"/g, '""')}"`;
}
function shouldApplyCsvFormulaWrap(value) {
  if (!/[0-9]$/.test(value)) {
    return false;
  }
  if (/^=".*"$/.test(value)) {
    return false;
  }
  if (/[/\-:.Tt]/.test(value)) {
    return false;
  }
  return true;
}
function getCsvDelimiterChar(delimiter) {
  if (delimiter === X2T_CSV_DELIMITER_TAB) return "	";
  if (delimiter === X2T_CSV_DELIMITER_SEMICOLON) return ";";
  return ",";
}
function isMultilineCsv(buffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  const delimiter = getCsvDelimiterChar(
    detectX2tCsvDelimiter(buffer, csvEncoding)
  );
  const text = decodeCsvBuffer(buffer, csvEncoding);
  const physicalLines = text.split(/\r?\n/).filter((line) => line.length > 0).length;
  const logicalRows = parseCsvText(text, delimiter).length;
  return physicalLines > logicalRows + 2;
}
function parseCsvBuffer(buffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  const delimiter = getCsvDelimiterChar(
    detectX2tCsvDelimiter(buffer, csvEncoding)
  );
  return parseCsvText(decodeCsvBuffer(buffer, csvEncoding), delimiter);
}
function sanitizeCsvBufferForX2t(buffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  const csvDelimiter = detectX2tCsvDelimiter(buffer, csvEncoding);
  const delimiter = getCsvDelimiterChar(csvDelimiter);
  const bytes = new Uint8Array(buffer);
  const withUtf8Bom = csvEncoding === X2T_CSV_ENCODING_UTF8 && bytes.length >= 3 && bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191;
  const text = decodeCsvBuffer(buffer, csvEncoding);
  const lineEnding = text.includes("\r\n") ? "\r\n" : "\n";
  const rows = parseCsvText(text, delimiter);
  const sanitized = rows.map((fields) => {
    for (let i = 0; i < fields.length - 1; i++) {
      if (shouldApplyCsvFormulaWrap(fields[i])) {
        fields[i] = toCsvFormulaCell(fields[i]);
      }
    }
    return serializeCsvRow(fields, delimiter);
  });
  return encodeCsvBuffer(sanitized.join(lineEnding), withUtf8Bom);
}
function getFirstCsvLine(text) {
  const newline = text.search(/\r?\n/);
  return newline === -1 ? text : text.slice(0, newline);
}
function detectX2tCsvDelimiter(buffer, encoding) {
  const line = getFirstCsvLine(decodeCsvSample(buffer, encoding));
  const counts = { comma: 0, semicolon: 0, tab: 0 };
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        i++;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (inQuotes) continue;
    if (ch === ",") counts.comma++;
    else if (ch === ";") counts.semicolon++;
    else if (ch === "	") counts.tab++;
  }
  if (counts.tab > counts.comma && counts.tab > counts.semicolon) {
    return X2T_CSV_DELIMITER_TAB;
  }
  if (counts.semicolon > counts.comma) {
    return X2T_CSV_DELIMITER_SEMICOLON;
  }
  return X2T_CSV_DELIMITER_COMMA;
}
function detectX2tCsvEncoding(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes.length >= 3 && bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191) {
    return X2T_CSV_ENCODING_UTF8;
  }
  return isValidUtf8(bytes) ? X2T_CSV_ENCODING_UTF8 : X2T_CSV_ENCODING_GBK;
}
function getX2tCsvConvertOptions(buffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  return {
    csvEncoding,
    csvDelimiter: detectX2tCsvDelimiter(buffer, csvEncoding)
  };
}
function getX2tExportFormats(fileType, sourceFileType) {
  const ext = normalizeX2tExportFileType(fileType);
  const source = sourceFileType ?? fileType;
  const formatTo = x2tSourceFormatByExt[ext] ?? getDefaultX2tSourceFormat(source);
  return {
    formatFrom: getX2tExportBinFormat(source),
    formatTo
  };
}
function normalizeX2tExportFileType(fileType) {
  const ext = getFileExt(fileType);
  return ext === "doc" ? "docx" : ext;
}
var outputFormatToExt = Object.fromEntries(
  Object.entries(x2tSourceFormatByExt).map(([ext, format]) => [format, ext])
);
Object.assign(outputFormatToExt, {
  [70 /* AVS_FILE_DOCUMENT_HTML */]: "html",
  [71 /* AVS_FILE_DOCUMENT_MHT */]: "mht",
  [72 /* AVS_FILE_DOCUMENT_EPUB */]: "epub",
  [73 /* AVS_FILE_DOCUMENT_FB2 */]: "fb2",
  [74 /* AVS_FILE_DOCUMENT_MOBI */]: "mobi",
  [76 /* AVS_FILE_DOCUMENT_DOTX */]: "dotx",
  [77 /* AVS_FILE_DOCUMENT_DOTM */]: "dotm",
  [78 /* AVS_FILE_DOCUMENT_ODT_FLAT */]: "fodt",
  [79 /* AVS_FILE_DOCUMENT_OTT */]: "ott",
  [85 /* AVS_FILE_DOCUMENT_OFORM */]: "oform",
  [86 /* AVS_FILE_DOCUMENT_DOCXF */]: "docxf",
  [132 /* AVS_FILE_PRESENTATION_PPSX */]: "ppsx",
  [133 /* AVS_FILE_PRESENTATION_PPTM */]: "pptm",
  [134 /* AVS_FILE_PRESENTATION_PPSM */]: "ppsm",
  [135 /* AVS_FILE_PRESENTATION_POTX */]: "potx",
  [136 /* AVS_FILE_PRESENTATION_POTM */]: "potm",
  [137 /* AVS_FILE_PRESENTATION_ODP_FLAT */]: "fodp",
  [138 /* AVS_FILE_PRESENTATION_OTP */]: "otp",
  [261 /* AVS_FILE_SPREADSHEET_XLSM */]: "xlsm",
  [262 /* AVS_FILE_SPREADSHEET_XLTX */]: "xltx",
  [263 /* AVS_FILE_SPREADSHEET_XLTM */]: "xltm",
  [264 /* AVS_FILE_SPREADSHEET_XLSB */]: "xlsb",
  [265 /* AVS_FILE_SPREADSHEET_ODS_FLAT */]: "fods",
  [266 /* AVS_FILE_SPREADSHEET_OTS */]: "ots",
  [513 /* AVS_FILE_CROSSPLATFORM_PDF */]: "pdf",
  [521 /* AVS_FILE_CROSSPLATFORM_PDFA */]: "pdf",
  [515 /* AVS_FILE_CROSSPLATFORM_DJVU */]: "djvu",
  [516 /* AVS_FILE_CROSSPLATFORM_XPS */]: "xps",
  [2048 /* AVS_FILE_OTHER */ + 9]: "zip",
  [2056 /* AVS_FILE_OTHER_JSON */]: "json",
  [8193 /* AVS_FILE_CANVAS_WORD */]: "bin",
  [8194 /* AVS_FILE_CANVAS_SPREADSHEET */]: "bin",
  [8195 /* AVS_FILE_CANVAS_PRESENTATION */]: "bin",
  [16385 /* AVS_FILE_DRAW_VSDX */]: "vsdx",
  [16386 /* AVS_FILE_DRAW_VSSX */]: "vssx",
  [16387 /* AVS_FILE_DRAW_VSTX */]: "vstx",
  [16388 /* AVS_FILE_DRAW_VSDM */]: "vsdm",
  [16389 /* AVS_FILE_DRAW_VSSM */]: "vssm",
  [16390 /* AVS_FILE_DRAW_VSTM */]: "vstm",
  [4097 /* AVS_FILE_TEAMLAB_DOCY */]: "bin",
  [4098 /* AVS_FILE_TEAMLAB_XLSY */]: "bin",
  [4099 /* AVS_FILE_TEAMLAB_PPTY */]: "bin"
});
var canvasBinOutputFormats = /* @__PURE__ */ new Set([
  8193 /* AVS_FILE_CANVAS_WORD */,
  8194 /* AVS_FILE_CANVAS_SPREADSHEET */,
  8195 /* AVS_FILE_CANVAS_PRESENTATION */,
  4097 /* AVS_FILE_TEAMLAB_DOCY */,
  4098 /* AVS_FILE_TEAMLAB_XLSY */,
  4099 /* AVS_FILE_TEAMLAB_PPTY */
]);
function isCanvasBinOutputFormat(outputFormat) {
  return outputFormat != null && canvasBinOutputFormats.has(outputFormat);
}
function extensionFromOutputFormat(outputFormat) {
  if (outputFormat == null) {
    return "";
  }
  return outputFormatToExt[outputFormat] || "";
}
function ensureTitleWithExtension(title, fileType) {
  const ext = fileType.toLowerCase();
  if (!ext) {
    return title;
  }
  const currentExt = getFileExt(title);
  if (currentExt === ext) {
    return title;
  }
  const dotIndex = title.lastIndexOf(".");
  const base = dotIndex > 0 && dotIndex < title.length - 1 ? title.slice(0, dotIndex) : title;
  return `${base}.${ext}`;
}

// vendor/onlyoffice-web-comp/store/document.ts
var DEFAULT_DOCUMENT_SCOPE = "__default__";
function createDefaultDocumentState() {
  return {
    isNew: true,
    fileName: "New Document.docx",
    fileType: "docx"
  };
}
function normalizeDocumentState(state) {
  return {
    ...state,
    isNew: state.isNew ?? !state.file,
    fileType: state.fileType || getFileExt(state.fileName) || "docx"
  };
}
var documentStates = /* @__PURE__ */ new Map();
documentStates.set(DEFAULT_DOCUMENT_SCOPE, createDefaultDocumentState());
function getScopeId(scopeId) {
  return scopeId || DEFAULT_DOCUMENT_SCOPE;
}
function setDocumentObj(state, scopeId) {
  documentStates.set(getScopeId(scopeId), normalizeDocumentState(state));
}
function getDocumentObj(scopeId) {
  const key = getScopeId(scopeId);
  let state = documentStates.get(key);
  if (!state) {
    state = createDefaultDocumentState();
    documentStates.set(key, state);
  }
  return state;
}
function clearDocumentObj(scopeId) {
  if (scopeId) {
    documentStates.delete(scopeId);
    return;
  }
  documentStates.set(DEFAULT_DOCUMENT_SCOPE, createDefaultDocumentState());
}

// vendor/onlyoffice-web-comp/store/lang.ts
var currentLang = ONLYOFFICE_LANG_KEY.ZH;
function getCurrentLang() {
  return currentLang;
}
function setCurrentLang(lang) {
  currentLang = lang;
}
function getOnlyOfficeLang() {
  return getCurrentLang();
}

// vendor/onlyoffice-web-comp/util/initialize.ts
var initializePromise = null;
var initializeApiUrl = "";
function resetLoadedDocsApi(apiUrl) {
  if (!initializeApiUrl || initializeApiUrl === apiUrl) {
    return;
  }
  document.querySelectorAll(
    'script[src*="/web-apps/apps/api/documents/api.js"]'
  ).forEach((script) => script.remove());
  document.querySelectorAll("iframe[data-onlyoffice-preload]").forEach((iframe) => iframe.remove());
  try {
    delete window.DocsAPI;
  } catch {
    window.DocsAPI = void 0;
  }
}
function preloadEditorFrame() {
  if (document.querySelector(`iframe[data-onlyoffice-preload="${STATIC_RESOURCE.onlyoffice.preloadHtml}"]`)) {
    return;
  }
  const iframe = document.createElement("iframe");
  iframe.src = STATIC_RESOURCE.onlyoffice.preloadUrl;
  iframe.dataset.onlyofficePreload = STATIC_RESOURCE.onlyoffice.preloadHtml;
  iframe.style.position = "absolute";
  iframe.style.width = "0";
  iframe.style.height = "0";
  iframe.style.border = "0";
  iframe.style.left = "-9999px";
  iframe.style.top = "-9999px";
  iframe.style.visibility = "hidden";
  iframe.style.pointerEvents = "none";
  iframe.setAttribute("aria-hidden", "true");
  iframe.setAttribute("tabindex", "-1");
  document.body.appendChild(iframe);
}
async function initializeOnlyOffice() {
  if (typeof window === "undefined") return;
  const apiUrl = STATIC_RESOURCE.onlyoffice.apiUrl;
  if (initializePromise && initializeApiUrl === apiUrl) {
    return initializePromise;
  }
  if (initializePromise && initializeApiUrl !== apiUrl) {
    initializePromise = null;
  }
  resetLoadedDocsApi(apiUrl);
  initializeApiUrl = apiUrl;
  initializePromise = new Promise((resolve, reject) => {
    preloadEditorFrame();
    if (window.DocsAPI?.DocEditor) {
      resolve();
      return;
    }
    let script = document.querySelector(
      `script[src="${apiUrl}"]`
    );
    if (!script) {
      script = document.createElement("script");
      script.src = apiUrl;
      document.head.appendChild(script);
    }
    script.addEventListener("load", () => resolve(), { once: true });
    script.addEventListener(
      "error",
      () => {
        initializePromise = null;
        reject(new Error("Failed to load OnlyOffice DocsAPI script"));
      },
      { once: true }
    );
  });
  return initializePromise;
}

// vendor/onlyoffice-web-comp/internal/editor/x2t.ts
var workerFactory = null;
function setX2tWorkerFactory(factory) {
  workerFactory = factory;
}
function createX2tWorker() {
  if (!workerFactory) {
    throw new Error(
      "x2t worker factory is not registered (call setX2tWorkerFactory before opening documents)"
    );
  }
  return workerFactory();
}
var X2tConverter = class {
  worker = null;
  initPromise = null;
  messageId = 0;
  pendingMessages = /* @__PURE__ */ new Map();
  resourceKey = "";
  logger;
  getWorkerStaticResource() {
    const staticResource = getStaticResource();
    if (typeof window === "undefined") {
      return staticResource;
    }
    const origin = window.location.origin;
    return {
      ...staticResource,
      version: { ...staticResource.version },
      onlyoffice: { ...staticResource.onlyoffice },
      x2t: {
        ...staticResource.x2t,
        root: resolveSiteUrl(origin, staticResource.x2t.root),
        script: resolveSiteUrl(origin, staticResource.x2t.script),
        wasm: resolveSiteUrl(origin, staticResource.x2t.wasm),
        pdfFonts: {
          root: resolveSiteUrl(origin, staticResource.x2t.pdfFonts.root),
          default: resolveSiteUrl(origin, staticResource.x2t.pdfFonts.default)
        }
      }
    };
  }
  /**
   * @description 生成递增的 worker 消息 ID。
   */
  getNextId() {
    return ++this.messageId;
  }
  logRaw(level, message, consoleArgs) {
    if (this.logger) {
      this.logger.raw(level, "worker", message, consoleArgs);
      return;
    }
    console[level](...consoleArgs);
  }
  /**
   * @description 向 worker 发送请求并等待对应响应。
   */
  sendMessage(type, payload) {
    return new Promise((resolve, reject) => {
      if (!this.worker) {
        reject(new Error("Worker not initialized"));
        return;
      }
      const id = this.getNextId();
      this.pendingMessages.set(id, { resolve, reject });
      if (type === "convert" && payload?.data instanceof ArrayBuffer) {
        this.worker.postMessage({ id, type, payload }, [payload.data]);
      } else {
        this.worker.postMessage({ id, type, payload });
      }
    });
  }
  /**
   * @description 处理 worker 返回的响应消息。
   */
  handleWorkerMessage = (event) => {
    const { id, type, payload, error, errorDetails } = event.data;
    if (type === "ready") {
      this.logRaw("log", "worker ready", ["[X2tConverter] Worker ready"]);
      return;
    }
    const pending = this.pendingMessages.get(id);
    if (!pending) return;
    this.pendingMessages.delete(id);
    if (type === "error") {
      const errorMessage = error || "Unknown worker error";
      const details = errorDetails && typeof errorDetails === "object" ? { message: errorMessage, ...errorDetails } : { message: errorMessage };
      if (this.logger) {
        this.logger.error("worker", "worker request failed", details);
      } else {
        console.error("[X2tConverter] Worker request failed:", details);
      }
      pending.reject(new Error(errorMessage));
    } else {
      pending.resolve(payload);
    }
  };
  /**
   * @description 处理 worker 运行错误，并让所有等待中的请求失败。
   */
  handleWorkerError = (error) => {
    this.logRaw("error", "worker error", [
      "[X2tConverter] Worker error:",
      error
    ]);
    for (const [id, pending] of this.pendingMessages) {
      pending.reject(new Error(`Worker error: ${error.message}`));
      this.pendingMessages.delete(id);
    }
  };
  /**
   * @description 初始化 x2t worker；重复调用会复用同一个初始化 Promise。
   */
  init(logger) {
    if (logger) {
      this.logger = logger;
    }
    if (this.initPromise) {
      return this.initPromise;
    }
    this.initPromise = new Promise((resolve, reject) => {
      try {
        this.worker = createX2tWorker();
        this.worker.onmessage = this.handleWorkerMessage;
        this.worker.onerror = this.handleWorkerError;
        this.logRaw("log", "worker created", ["[X2tConverter] Worker created"]);
        resolve();
      } catch (err) {
        this.initPromise = null;
        reject(err);
      }
    });
    return this.initPromise;
  }
  /**
   * @description 将文档从一种格式转换为另一种格式。
   */
  async convert({
    data,
    fileFrom,
    fileTo,
    formatFrom,
    formatTo,
    media,
    pdfBin,
    fonts,
    fontAliases,
    fontExportAliases,
    themes,
    csvEncoding,
    csvDelimiter,
    csvDelimiterChar
  }, logger) {
    if (logger) {
      this.logger = logger;
    }
    const staticResource = this.getWorkerStaticResource();
    const resourceKey = JSON.stringify(staticResource.x2t);
    if (this.worker && this.resourceKey && this.resourceKey !== resourceKey) {
      this.terminate(logger);
    }
    this.resourceKey = resourceKey;
    await this.init(logger);
    const cloneMap = (map) => {
      if (!map) return void 0;
      return Object.fromEntries(
        Object.entries(map).map(([key, value]) => [key, value.slice(0)])
      );
    };
    const dataClone = data.slice(0);
    const payload = {
      data: dataClone,
      fileFrom,
      fileTo,
      formatFrom,
      formatTo,
      media: cloneMap(media),
      pdfBin: pdfBin?.slice(0),
      fonts: cloneMap(fonts),
      fontAliases,
      fontExportAliases,
      themes: cloneMap(themes),
      csvEncoding,
      csvDelimiter,
      csvDelimiterChar,
      staticResource
    };
    this.logger?.worker("convert", {
      fileFrom,
      fileTo,
      formatFrom,
      formatTo
    });
    return this.sendMessage("convert", payload);
  }
  /**
   * @description 终止 worker 并释放关联资源。
   */
  terminate(logger) {
    if (logger) {
      this.logger = logger;
    }
    if (this.worker) {
      for (const [id, pending] of this.pendingMessages) {
        pending.reject(new Error("Worker terminated"));
        this.pendingMessages.delete(id);
      }
      this.worker.terminate();
      this.worker = null;
      this.initPromise = null;
      this.logRaw("log", "worker terminated", [
        "[X2tConverter] Worker terminated"
      ]);
    }
  }
  /**
   * @description 判断 worker 是否已经初始化。
   */
  get isInitialized() {
    return this.worker !== null && this.initPromise !== null;
  }
};
var converter = new X2tConverter();

// vendor/onlyoffice-web-comp/util/event-emitter.ts
var EventEmitter = class {
  listeners = /* @__PURE__ */ new Map();
  on(event, listener) {
    const list = this.listeners.get(event) ?? [];
    list.push(listener);
    this.listeners.set(event, list);
    return this;
  }
  once(event, listener) {
    const wrapper = (...args) => {
      this.off(event, wrapper);
      listener(...args);
    };
    return this.on(event, wrapper);
  }
  off(event, listener) {
    if (!listener) {
      this.listeners.delete(event);
      return this;
    }
    const list = this.listeners.get(event);
    if (!list) {
      return this;
    }
    const next = list.filter((fn) => fn !== listener);
    if (next.length) {
      this.listeners.set(event, next);
    } else {
      this.listeners.delete(event);
    }
    return this;
  }
  removeAllListeners(event) {
    if (event === void 0) {
      this.listeners.clear();
    } else {
      this.listeners.delete(event);
    }
    return this;
  }
  emit(event, ...args) {
    const list = this.listeners.get(event);
    if (!list?.length) {
      return false;
    }
    for (const listener of [...list]) {
      listener(...args);
    }
    return true;
  }
};

// vendor/onlyoffice-web-comp/internal/editor/runtime-bridge.ts
var MockSocket = class _MockSocket {
  static _staticEmitter = new EventEmitter();
  static on(event, listener) {
    _MockSocket._staticEmitter.on(event, listener);
  }
  static off(event, listener) {
    _MockSocket._staticEmitter.off(event, listener);
  }
  active = true;
  connected = false;
  disconnected = true;
  recovered = false;
  id = "";
  io = {
    setOpenToken: () => {
    },
    setSessionToken: () => {
    },
    on: function() {
      return this;
    },
    reconnectionAttempts: function() {
      return this;
    },
    reconnectionDelay: function() {
      return this;
    },
    reconnectionDelayMax: function() {
      return this;
    },
    timeout: function() {
      return this;
    },
    transports: function() {
      return this;
    },
    upgrade: function() {
      return this;
    },
    upgradeTransport: function() {
      return this;
    },
    upgradeTimeout: function() {
      return this;
    }
  };
  _clientEmitter = new EventEmitter();
  _serverEmitter = new EventEmitter();
  _debug;
  _logger;
  constructor(options = {}) {
    this._debug = options.debug;
    this._logger = options.logger;
    if (!options.deferConnect) {
      this.connect();
    }
  }
  _log(...args) {
    if (this._logger) {
      this._logger.raw("log", "socket", "mock socket", ["[MockSocket]", ...args]);
      return;
    }
    if (this._debug) {
      console.log("[MockSocket]", ...args);
    }
  }
  open() {
    return this.connect();
  }
  compress() {
  }
  /**
   * @description 模拟连接建立并生成新的 session id。
   */
  connect() {
    this.connected = true;
    this.disconnected = false;
    this.id = Math.random().toString(36).substring(2, 15);
    setTimeout(() => {
      this._trigger("connect");
      _MockSocket._staticEmitter.emit("connect", { socket: this });
    }, 0);
    return this;
  }
  disconnect() {
    this.connected = false;
    this.disconnected = true;
    this._trigger("disconnect");
    _MockSocket._staticEmitter.emit("disconnect", { socket: this });
    return this;
  }
  close() {
    return this.disconnect();
  }
  /**
   * @description 触发本地监听器，用于模拟服务端下行事件。
   */
  _trigger(event, ...args) {
    this._log(`trigger event: ${event}`, ...args);
    this._clientEmitter.emit(event, ...args);
    return this;
  }
  /**
   * @description 注册服务端下行事件监听器。
   */
  on(event, listener) {
    this._clientEmitter.on(event, listener);
    return this;
  }
  /**
   * @description 注册一次性的服务端下行事件监听器。
   */
  once(event, listener) {
    this._clientEmitter.once(event, listener);
    return this;
  }
  /**
   * @description 移除事件监听器。
   */
  off(event, listener) {
    this._clientEmitter.off(event, listener);
    return this;
  }
  /**
   * @description 移除全部监听器，或移除指定事件的监听器。
   */
  removeAllListeners(event) {
    this._clientEmitter.removeAllListeners(event);
    return this;
  }
  /**
   * @description 使用 message 事件向服务端发送消息。
   */
  send(...args) {
    if (!this.connected) return this;
    this.emit("message", ...args);
    return this;
  }
  /**
   * @description 向服务端发送事件消息。
   */
  emit(event, ...args) {
    this._log(`emit: ${event}`, ...args);
    if (!this.connected) return this;
    const processEmit = async () => {
      this._serverEmitter.emit(event, ...args);
    };
    setTimeout(() => processEmit(), 0);
    return this;
  }
  server = {
    on: (event, listener) => {
      this._serverEmitter.on(event, listener);
    },
    off: (event, listener) => {
      this._serverEmitter.off(event, listener);
    },
    emit: (event, ...args) => {
      this._clientEmitter.emit(event, ...args);
    }
  };
};
function io(_url, options) {
  return new MockSocket(options);
}
var ioWithStatics = io;
var runtime_bridge_default = ioWithStatics;
function isForbiddenRequestHeader(name) {
  const lowerName = name.toLowerCase();
  return lowerName === "accept-charset" || lowerName === "accept-encoding" || lowerName === "access-control-request-headers" || lowerName === "access-control-request-method" || lowerName === "connection" || lowerName === "content-length" || lowerName === "cookie" || lowerName === "cookie2" || lowerName === "date" || lowerName === "dnt" || lowerName === "expect" || lowerName === "host" || lowerName === "keep-alive" || lowerName === "origin" || lowerName === "referer" || lowerName === "te" || lowerName === "trailer" || lowerName === "transfer-encoding" || lowerName === "upgrade" || lowerName === "via" || lowerName.startsWith("proxy-") || lowerName.startsWith("sec-");
}
function createXHRProxy(BaseXHR = globalThis.XMLHttpRequest, options = {}) {
  return class ProxyXMLHttpRequest extends BaseXHR {
    static _middlewares = [];
    _isMocked = false;
    _requestMethod = "GET";
    _requestUrl = "";
    _requestHeaders = new Headers();
    _requestBody = null;
    _responseHeaders = new Headers();
    /**
     * @description 注册全局 XHR 中间件。
     */
    static use(middleware) {
      this._middlewares.push(middleware);
    }
    /**
     * @description 清空全部 XHR 中间件。
     */
    static clearMiddlewares() {
      this._middlewares = [];
    }
    open(method, url, async = true, username, password) {
      const normalizedUrl = (() => {
        try {
          return options.baseUrl ? new URL(url.toString(), options.baseUrl).href : url.toString();
        } catch {
          return url.toString();
        }
      })();
      this._requestMethod = method;
      this._requestUrl = normalizedUrl;
      this._requestHeaders = new Headers();
      this._responseHeaders = new Headers();
      this._isMocked = false;
      super.open(
        method,
        normalizedUrl,
        async,
        username ?? void 0,
        password ?? void 0
      );
    }
    setRequestHeader(name, value) {
      if (isForbiddenRequestHeader(name)) {
        return;
      }
      this._requestHeaders.append(name, value);
      if (!this._isMocked) {
        super.setRequestHeader(name, value);
      }
    }
    getResponseHeader(name) {
      if (this._isMocked) {
        return this._responseHeaders.get(name);
      }
      return super.getResponseHeader(name);
    }
    getAllResponseHeaders() {
      if (!this._isMocked) {
        return super.getAllResponseHeaders();
      }
      return Array.from(this._responseHeaders.entries()).map(([key, value]) => `${key}: ${value}`).join("\r\n");
    }
    send(body) {
      this._requestBody = body;
      if (options.shouldBypass?.(this._requestUrl, this._requestMethod)) {
        super.send(body);
        return;
      }
      this._tryMiddlewares().then((handled) => {
        if (!handled) {
          super.send(body);
        }
      }).catch((err) => {
        console.error("ProxyXMLHttpRequest middleware error:", err);
        super.send(body);
      });
    }
    async _tryMiddlewares() {
      let request;
      try {
        const reqInit = {
          method: this._requestMethod,
          headers: this._requestHeaders,
          body: this._requestBody,
          mode: "cors"
        };
        if (this.withCredentials) {
          reqInit.credentials = "include";
        }
        request = new Request(this._requestUrl, reqInit);
        console.log("ProxyXHR created request:", {
          url: this._requestUrl,
          method: request.method,
          hasBody: !!request.body,
          originalBody: this._requestBody
        });
      } catch (e) {
        return false;
      }
      for (const mw of ProxyXMLHttpRequest._middlewares) {
        const response = await mw(request.clone());
        if (response) {
          this._isMocked = true;
          await this._handleMockResponse(response);
          return true;
        }
      }
      return false;
    }
    async _handleMockResponse(response) {
      this._responseHeaders = new Headers(response.headers);
      const emit = (event) => {
        this.dispatchEvent(event);
      };
      emit(new ProgressEvent("loadstart"));
      Object.defineProperty(this, "readyState", {
        value: 2,
        writable: false,
        configurable: true
      });
      emit(new Event("readystatechange"));
      Object.defineProperty(this, "readyState", {
        value: 3,
        writable: false,
        configurable: true
      });
      emit(new Event("readystatechange"));
      try {
        let responseData;
        if (this.responseType === "json") {
          responseData = await response.json();
        } else if (this.responseType === "arraybuffer") {
          responseData = await response.arrayBuffer();
        } else if (this.responseType === "blob") {
          responseData = await response.blob();
        } else if (this.responseType === "document") {
          const text = await response.text();
          responseData = new DOMParser().parseFromString(text, "text/xml");
        } else {
          responseData = await response.text();
        }
        Object.defineProperty(this, "status", {
          value: response.status,
          writable: false,
          configurable: true
        });
        Object.defineProperty(this, "statusText", {
          value: response.statusText,
          writable: false,
          configurable: true
        });
        Object.defineProperty(this, "response", {
          value: responseData,
          writable: false,
          configurable: true
        });
        Object.defineProperty(this, "responseText", {
          value: typeof responseData === "string" ? responseData : JSON.stringify(responseData),
          writable: false,
          configurable: true
        });
        Object.defineProperty(this, "responseURL", {
          value: response.url,
          writable: false,
          configurable: true
        });
        emit(
          new ProgressEvent("progress", {
            lengthComputable: true,
            loaded: 100,
            total: 100
          })
        );
        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true
        });
        emit(new Event("readystatechange"));
        emit(new ProgressEvent("load"));
        emit(new ProgressEvent("loadend"));
      } catch (e) {
        console.error("ProxyXHR: error handling response", e);
        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true
        });
        emit(new Event("readystatechange"));
        emit(new ProgressEvent("error"));
        emit(new ProgressEvent("loadend"));
      }
    }
  };
}
function createFetchProxy(target = globalThis.fetch, options = {}) {
  const middlewares = [];
  const BaseFetch = typeof target === "function" ? target : target.fetch.bind(target);
  const proxy = (async (input, init) => {
    let request;
    try {
      const normalizedInput = options.baseUrl && !(input instanceof Request) ? new URL(input.toString(), options.baseUrl).href : input;
      request = new Request(normalizedInput, init);
    } catch (e) {
      return BaseFetch(input, init);
    }
    try {
      for (const mw of middlewares) {
        const response = await mw(request.clone());
        if (response) {
          return response;
        }
      }
    } catch (err) {
      console.error("ProxyFetch middleware error:", err);
      return BaseFetch(request);
    }
    return BaseFetch(request);
  });
  proxy.use = (middleware) => {
    middlewares.push(middleware);
  };
  proxy.clearMiddlewares = () => {
    middlewares.length = 0;
  };
  return proxy;
}
var CROSS_ORIGIN_BRIDGE_MESSAGE = {
  EDITOR_COMMAND: "editor:command",
  EDITOR_RESPONSE: "editor:response",
  EDITOR_EVENT: "editor:event",
  EDITOR_SET_READONLY: "editor:set-readonly"
};
var CROSS_ORIGIN_EDITOR_COMMAND = {
  EDITOR_SUBSCRIBE: "editor:subscribe",
  DOCUMENT_RENAME: "document:rename",
  COMMENT_ADD: "comment:add",
  COMMENT_UPDATE: "comment:update",
  COMMENT_REMOVE: "comment:remove",
  COMMENT_GO_TO: "comment:go-to",
  COMMENT_LIST: "comment:list",
  COMMENT_SUBSCRIBE: "comment:subscribe",
  REVISION_ADD_DEMO: "revision:add-demo",
  REVISION_LIST: "revision:list",
  REVISION_SET_TRACK: "revision:set-track",
  REVISION_IS_TRACK: "revision:is-track",
  REVISION_HAVE_CHANGES: "revision:have-changes",
  REVISION_PREPARE_REVIEW: "revision:prepare-review",
  REVISION_NEXT: "revision:next",
  REVISION_PREV: "revision:prev",
  REVISION_GO_TO: "revision:go-to",
  REVISION_ACCEPT: "revision:accept",
  REVISION_REJECT: "revision:reject",
  REVISION_ACCEPT_ALL: "revision:accept-all",
  REVISION_REJECT_ALL: "revision:reject-all",
  REVISION_ACCEPT_SELECTION: "revision:accept-selection",
  REVISION_REJECT_SELECTION: "revision:reject-selection",
  REVISION_SUBSCRIBE: "revision:subscribe"
};
var CROSS_ORIGIN_EDITOR_EVENT = {
  ADD_COMMENT: "asc_onAddComment",
  CHANGE_COMMENT: "asc_onChangeCommentData",
  REMOVE_COMMENT: "asc_onRemoveComment",
  SHOW_REVISIONS_CHANGE: "asc_onShowRevisionsChange",
  TRACK_REVISIONS_CHANGE: "asc_onOnTrackRevisionsChange",
  DOCUMENT_MODIFIED_CHANGED: "asc_onDocumentModifiedChanged"
};
function shouldBypassOnlyOfficeProxy(url, baseUrl) {
  const pathname = new URL(url, baseUrl).pathname;
  return pathname.includes("/sdkjs/common/AllFonts.js") || pathname.includes("/sdkjs/common/libfont/") || pathname.includes("/fonts/");
}
function getScopedIoRegistry(win = window) {
  const parent = win;
  if (!parent.__ONLYOFFICE_SCOPED_IO__) {
    parent.__ONLYOFFICE_SCOPED_IO__ = {};
  }
  return parent.__ONLYOFFICE_SCOPED_IO__;
}
function registerScopedIo(containerId, factory, win = window) {
  getScopedIoRegistry(win)[containerId] = factory;
}
function unregisterScopedIo(containerId, win = window) {
  const registry = win.__ONLYOFFICE_SCOPED_IO__;
  if (registry) {
    delete registry[containerId];
  }
}
function extractDownloadFileName(url) {
  if (!url) {
    return "download";
  }
  try {
    const parsed = new URL(url, window.location.href);
    const fromQuery = parsed.searchParams.get("filename");
    if (fromQuery) {
      return decodeURIComponent(fromQuery);
    }
    const pathname = parsed.pathname;
    const name = decodeURIComponent(pathname.split("/").pop() || "");
    if (name.startsWith("output.")) {
      return name;
    }
    if (name) {
      return name;
    }
  } catch {
    const fallback = url.split("/").pop() || url.split("?")[0];
    if (fallback) {
      return decodeURIComponent(fallback);
    }
  }
  return "download";
}
function parseContentDispositionFileName(header) {
  if (!header) {
    return "";
  }
  const utf8Match = /filename\*=UTF-8''([^;]+)/i.exec(header);
  if (utf8Match?.[1]) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch {
      return utf8Match[1];
    }
  }
  const asciiMatch = /filename="([^"]+)"/i.exec(header);
  if (asciiMatch?.[1]) {
    return asciiMatch[1];
  }
  return "";
}
function scheduleNamedDownloadPatch(win, server, retries = 100) {
  if (win.__ONLYOFFICE_GETFILE_PATCHED__) {
    return;
  }
  if (win.AscCommon?.getFile) {
    installNamedDownloadPatch(win, server);
    return;
  }
  if (retries > 0) {
    win.setTimeout(
      () => scheduleNamedDownloadPatch(win, server, retries - 1),
      50
    );
  }
}
function extractOutputNameFromCacheUrl(url) {
  const match = /\/cache\/files\/data\/[^/]+\/(output\.[^/?#]+)/i.exec(url);
  return match?.[1] ?? "";
}
function triggerBlobDownload(win, blob, fileName) {
  const objectUrl = URL.createObjectURL(blob);
  const link = win.document.createElement("a");
  link.href = objectUrl;
  link.download = fileName;
  link.style.display = "none";
  win.document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(objectUrl);
}
function installNamedDownloadPatch(win, server) {
  win.__ONLYOFFICE_PROXY_SERVER__ = server;
  const ascCommon = win.AscCommon;
  if (!ascCommon?.getFile || win.__ONLYOFFICE_GETFILE_PATCHED__) {
    return;
  }
  const nativeGetFile = ascCommon.getFile.bind(ascCommon);
  const fetchFile = win.fetch.bind(win);
  ascCommon.getFile = (url) => {
    if (typeof url !== "string" || !url) {
      nativeGetFile(url);
      return;
    }
    const needsNamedDownload = url.includes("/cache/files/") || url.startsWith("blob:");
    if (!needsNamedDownload) {
      nativeGetFile(url);
      return;
    }
    const fallbackName = extractDownloadFileName(url);
    void fetchFile(url).then((response) => {
      if (!response.ok) {
        throw new Error(`Download failed: ${response.status}`);
      }
      const fileName = parseContentDispositionFileName(
        response.headers.get("Content-Disposition")
      ) || fallbackName;
      return response.blob().then((blob) => ({ blob, fileName }));
    }).then(({ blob, fileName }) => {
      triggerBlobDownload(win, blob, fileName);
    }).catch((err) => {
      console.warn("[OnlyOffice] named download fetch failed:", err);
      const outputName = extractOutputNameFromCacheUrl(url);
      const currentServer = win.__ONLYOFFICE_PROXY_SERVER__ ?? server;
      const blobUrl = outputName ? currentServer.getStoredOutputUrl(outputName) : null;
      const fileName = outputName && currentServer.getStoredOutputFileName(outputName) || fallbackName;
      if (blobUrl) {
        void fetchFile(blobUrl).then((response) => response.blob()).then((blob) => triggerBlobDownload(win, blob, fileName)).catch((fallbackErr) => {
          console.warn("[OnlyOffice] blob download fallback:", fallbackErr);
          nativeGetFile(url);
        });
        return;
      }
      nativeGetFile(url);
    });
  };
  win.__ONLYOFFICE_GETFILE_PATCHED__ = true;
}
function installPrintFramePatch(win, server) {
  if (win.__ONLYOFFICE_PRINT_FRAME_PATCHED__) {
    return;
  }
  const iframePrototype = win.HTMLIFrameElement?.prototype;
  const srcDescriptor = iframePrototype ? Object.getOwnPropertyDescriptor(iframePrototype, "src") : void 0;
  if (!srcDescriptor?.set || !srcDescriptor.get) {
    return;
  }
  const nativeSetter = srcDescriptor.set;
  const nativeGetter = srcDescriptor.get;
  const fetchFile = win.fetch.bind(win);
  Object.defineProperty(iframePrototype, "src", {
    configurable: srcDescriptor.configurable,
    enumerable: srcDescriptor.enumerable,
    get: nativeGetter,
    set(value) {
      const outputName = extractOutputNameFromCacheUrl(value);
      const blobUrl = outputName ? server.getStoredOutputUrl(outputName) : null;
      if (!blobUrl || !value.includes("/cache/files/")) {
        nativeSetter.call(this, value);
        return;
      }
      void fetchFile(blobUrl).then((response) => {
        if (!response.ok) {
          throw new Error(`Print PDF fetch failed: ${response.status}`);
        }
        return response.blob();
      }).then((blob) => {
        const objectUrl = win.URL.createObjectURL(blob);
        nativeSetter.call(this, objectUrl);
        win.setTimeout(() => win.URL.revokeObjectURL(objectUrl), 6e4);
      }).catch((error) => {
        console.warn("[OnlyOffice] print PDF fetch failed:", error);
        nativeSetter.call(this, value);
      });
    }
  });
  win.__ONLYOFFICE_PRINT_FRAME_PATCHED__ = true;
}
function installOnlyOfficeProxies(win, server, createIo, options = {}) {
  win.__ONLYOFFICE_PROXY_SERVER__ = server;
  if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
    scheduleNamedDownloadPatch(win, server);
    installPrintFramePatch(win, server);
    return;
  }
  const xhr = createXHRProxy(win.XMLHttpRequest, {
    baseUrl: win.location.href,
    shouldBypass: (url) => shouldBypassOnlyOfficeProxy(url, win.location.href)
  });
  const fetchProxy = createFetchProxy(win, { baseUrl: win.location.href });
  const WorkerCtor = win.Worker;
  xhr.use((request) => win.__ONLYOFFICE_PROXY_SERVER__?.handleRequest(request) ?? null);
  fetchProxy.use(
    (request) => win.__ONLYOFFICE_PROXY_SERVER__?.handleRequest(request) ?? null
  );
  const patches = {
    XMLHttpRequest: xhr,
    fetch: fetchProxy,
    Worker: function Worker2(url, options2) {
      const u = new URL(url, win.location.origin);
      return new WorkerCtor(
        u.href.replace(u.origin, win.location.origin),
        options2
      );
    }
  };
  if (options.installIo !== false) {
    patches.io = createIo;
  }
  Object.assign(win, patches);
  win.__ONLYOFFICE_PROXIES_INSTALLED__ = true;
  scheduleNamedDownloadPatch(win, server);
  installPrintFramePatch(win, server);
}
var REPORTER_HTML = "index.reporter.html";
function installReporterWindowHook(win, installProxies) {
  if (win.__ONLYOFFICE_REPORTER_HOOK__) {
    return;
  }
  win.__ONLYOFFICE_REPORTER_BRIDGE__ = { install: installProxies };
  win.__ONLYOFFICE_REPORTER_HOOK__ = true;
  const nativeOpen = win.open.bind(win);
  win.open = function openReporter(url, target, features) {
    const popup = nativeOpen(url, target, features);
    const href = typeof url === "string" ? url : url?.toString() ?? "";
    if (popup && href.includes(REPORTER_HTML)) {
      watchReporterWindow(popup, installProxies);
    }
    return popup;
  };
}
function watchReporterWindow(popup, installProxies) {
  const tryInstall = () => {
    if (popup.closed) {
      return true;
    }
    try {
      if (popup.location.href.includes(REPORTER_HTML)) {
        installProxies(popup);
        return true;
      }
    } catch {
    }
    return false;
  };
  if (tryInstall()) {
    return;
  }
  const interval = window.setInterval(() => {
    if (tryInstall()) {
      window.clearInterval(interval);
    }
  }, 1);
  popup.addEventListener(
    "load",
    () => {
      tryInstall();
      window.clearInterval(interval);
    },
    { once: true }
  );
}
var BRIDGE_SOURCE = "onlyoffice-bridge";
var sessions = /* @__PURE__ */ new Map();
var pendingRequests = /* @__PURE__ */ new Map();
var pendingReadyWaiters = /* @__PURE__ */ new Map();
var editorEventSubscribers = /* @__PURE__ */ new Map();
var listenerInstalled = false;
function isBridgeMessage(data) {
  return typeof data === "object" && data !== null && data.source === BRIDGE_SOURCE;
}
function getTargetOrigin(iframe) {
  try {
    return new URL(iframe.src, window.location.href).origin;
  } catch {
    return "*";
  }
}
function postToIframe(session, message, targetWindow) {
  const target = targetWindow ?? session.iframe.contentWindow;
  if (!target) {
    return;
  }
  target.postMessage(
    { ...message, source: BRIDGE_SOURCE },
    session.targetOrigin
  );
}
function getIframeByWindow(source) {
  if (!source || !("postMessage" in source)) {
    return null;
  }
  const frames = document.querySelectorAll(
    'iframe[name="frameEditor"]'
  );
  for (const frame of frames) {
    if (frame.contentWindow === source) {
      return frame;
    }
  }
  return null;
}
function updateSessionIframe(session, iframe) {
  if (session.iframe === iframe) {
    return;
  }
  detachSocket(session);
  session.iframe = iframe;
  session.targetOrigin = getTargetOrigin(iframe);
  session.bridgeReady = false;
  session.handshakeSent = false;
}
function attachSocket(session) {
  if (session.socket) {
    return;
  }
  const socket = session.createIo();
  session.socket = socket;
  session.server.registerSocketTransport(socket);
  const nativeServerEmit = socket.server.emit.bind(socket.server);
  socket.server.emit = (event, ...args) => {
    postToIframe(session, {
      type: "socket:event",
      frameEditorId: session.frameEditorId,
      event,
      args
    });
    return nativeServerEmit(event, ...args);
  };
}
function detachSocket(session) {
  if (!session.socket) {
    return;
  }
  session.server.handleDisconnect({ socket: session.socket });
  session.socket = null;
  session.handshakeSent = false;
}
async function handleHttpRequest(session, message) {
  const requestId = message.requestId;
  if (!requestId || !message.url || !message.method) {
    return;
  }
  try {
    const init = {
      method: message.method,
      headers: message.headers
    };
    if (message.method !== "GET" && message.method !== "HEAD") {
      init.body = decodeRequestBody(message);
    }
    let response = await session.server.handleRequest(
      new Request(message.url, init)
    );
    if (!response) {
      response = await fetch(message.url, init);
    }
    const responseHeaders = {};
    response.headers.forEach((value, key) => {
      responseHeaders[key] = value;
    });
    let responseBody = null;
    if (message.responseType === "arraybuffer") {
      const buffer = await response.arrayBuffer();
      responseBody = arrayBufferToBase64(buffer);
    } else {
      responseBody = await response.text();
    }
    postToIframe(session, {
      type: "http:response",
      frameEditorId: session.frameEditorId,
      requestId,
      status: response.status,
      responseHeaders,
      responseBody,
      responseType: message.responseType
    });
  } catch (error) {
    postToIframe(session, {
      type: "http:response",
      frameEditorId: session.frameEditorId,
      requestId,
      status: 0,
      responseBody: null
    });
    console.error("[OnlyOfficeBridge] HTTP proxy failed:", error);
  }
}
function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}
function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}
function decodeRequestBody(message) {
  if (message.body == null) {
    return null;
  }
  if (message.bodyEncoding === "base64") {
    return base64ToArrayBuffer(message.body);
  }
  return message.body;
}
function describeBridgeState(frameEditorId) {
  const session = sessions.get(frameEditorId);
  if (!session) {
    return `OnlyOffice cross-origin bridge is not ready for ${frameEditorId}: session is not registered`;
  }
  return `OnlyOffice cross-origin bridge is not ready for ${frameEditorId}: iframe=${session.iframe.src || "(empty)"}`;
}
function resolveReadyWaiters(frameEditorId, session) {
  const waiters = pendingReadyWaiters.get(frameEditorId);
  if (!waiters) {
    return;
  }
  pendingReadyWaiters.delete(frameEditorId);
  waiters.forEach((waiter) => {
    window.clearTimeout(waiter.timer);
    waiter.resolve(session);
  });
}
function rejectReadyWaiters(frameEditorId, error) {
  const waiters = pendingReadyWaiters.get(frameEditorId);
  if (!waiters) {
    return;
  }
  pendingReadyWaiters.delete(frameEditorId);
  waiters.forEach((waiter) => {
    window.clearTimeout(waiter.timer);
    waiter.reject(error);
  });
}
function waitForBridgeReady(frameEditorId, timeout) {
  const session = sessions.get(frameEditorId);
  if (session?.bridgeReady) {
    return Promise.resolve(session);
  }
  return new Promise((resolve, reject) => {
    const timer = window.setTimeout(() => {
      const waiters2 = pendingReadyWaiters.get(frameEditorId);
      waiters2?.delete(waiter);
      if (waiters2?.size === 0) {
        pendingReadyWaiters.delete(frameEditorId);
      }
      reject(new Error(describeBridgeState(frameEditorId)));
    }, timeout);
    const waiter = { resolve, reject, timer };
    let waiters = pendingReadyWaiters.get(frameEditorId);
    if (!waiters) {
      waiters = /* @__PURE__ */ new Set();
      pendingReadyWaiters.set(frameEditorId, waiters);
    }
    waiters.add(waiter);
  });
}
function handleBridgeMessage(event) {
  if (!isBridgeMessage(event.data)) {
    return;
  }
  const message = event.data;
  if (message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_RESPONSE && message.requestId) {
    const pending = pendingRequests.get(message.requestId);
    if (pending) {
      window.clearTimeout(pending.timer);
      pendingRequests.delete(message.requestId);
      if (message.error) {
        pending.reject(new Error(message.error));
      } else {
        pending.resolve(message.result);
      }
    }
    return;
  }
  if (message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_EVENT && message.event) {
    const frameEditorId2 = message.frameEditorId;
    const frameSubscribers = frameEditorId2 ? editorEventSubscribers.get(frameEditorId2) : void 0;
    const subscribers = frameSubscribers?.get(message.event);
    subscribers?.forEach((handler) => {
      handler(message.args ?? []);
    });
    return;
  }
  const frameEditorId = message.frameEditorId;
  if (!frameEditorId) {
    return;
  }
  const session = sessions.get(frameEditorId);
  if (!session) {
    return;
  }
  const sourceIframe = getIframeByWindow(event.source);
  if (sourceIframe) {
    updateSessionIframe(session, sourceIframe);
  }
  switch (message.type) {
    case "hello": {
      session.bridgeReady = true;
      resolveReadyWaiters(frameEditorId, session);
      attachSocket(session);
      const targetWindow = event.source && "postMessage" in event.source ? event.source : void 0;
      postToIframe(session, { type: "hello:ack", frameEditorId }, targetWindow);
      if (!session.handshakeSent) {
        session.handshakeSent = true;
        setTimeout(() => {
          if (session.socket) {
            session.server.sendCoAuthoringHandshake(session.socket);
          }
        }, 0);
      }
      if (session.pendingReadOnly !== null) {
        postToIframe(session, {
          type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
          frameEditorId,
          readOnly: session.pendingReadOnly
        });
        session.pendingReadOnly = null;
      }
      break;
    }
    case "socket:emit": {
      if (!session.socket || !message.event) {
        break;
      }
      session.socket.emit(message.event, ...message.args ?? []);
      break;
    }
    case "http": {
      void handleHttpRequest(session, message);
      break;
    }
    default:
      break;
  }
}
function registerCrossOriginBridge(frameEditorId, iframe, server, createIo) {
  const existing = sessions.get(frameEditorId);
  if (existing) {
    existing.server = server;
    existing.createIo = createIo;
    updateSessionIframe(existing, iframe);
    return;
  }
  sessions.set(frameEditorId, {
    frameEditorId,
    server,
    createIo,
    iframe,
    targetOrigin: getTargetOrigin(iframe),
    socket: null,
    bridgeReady: false,
    handshakeSent: false,
    pendingReadOnly: null
  });
  if (!listenerInstalled) {
    window.addEventListener("message", handleBridgeMessage);
    listenerInstalled = true;
  }
}
function unregisterCrossOriginBridge(frameEditorId) {
  const session = sessions.get(frameEditorId);
  if (session) {
    detachSocket(session);
    sessions.delete(frameEditorId);
    rejectReadyWaiters(
      frameEditorId,
      new Error(
        `OnlyOffice cross-origin bridge was unregistered: ${frameEditorId}`
      )
    );
  }
  editorEventSubscribers.delete(frameEditorId);
}
function setCrossOriginReadOnly(frameEditorId, readOnly) {
  const session = sessions.get(frameEditorId);
  if (!session) {
    return false;
  }
  session.pendingReadOnly = readOnly;
  if (!session.bridgeReady) {
    return true;
  }
  postToIframe(session, {
    type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
    frameEditorId,
    readOnly
  });
  session.pendingReadOnly = null;
  return true;
}
function callCrossOriginEditor(frameEditorId, command, payload = {}, timeout = 5e3) {
  const requestId = `editor-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  return waitForBridgeReady(frameEditorId, timeout).then(
    (session) => new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        pendingRequests.delete(requestId);
        reject(
          new Error(`OnlyOffice cross-origin command timed out: ${command}`)
        );
      }, timeout);
      pendingRequests.set(requestId, { resolve, reject, timer });
      postToIframe(session, {
        type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_COMMAND,
        frameEditorId,
        requestId,
        command,
        payload
      });
    })
  );
}
function subscribeCrossOriginEditorEvent(frameEditorId, event, handler) {
  let frameSubscribers = editorEventSubscribers.get(frameEditorId);
  if (!frameSubscribers) {
    frameSubscribers = /* @__PURE__ */ new Map();
    editorEventSubscribers.set(frameEditorId, frameSubscribers);
  }
  let subscribers = frameSubscribers.get(event);
  if (!subscribers) {
    subscribers = /* @__PURE__ */ new Set();
    frameSubscribers.set(event, subscribers);
  }
  subscribers.add(handler);
  return () => {
    subscribers?.delete(handler);
    if (subscribers?.size === 0) {
      frameSubscribers?.delete(event);
    }
    if (frameSubscribers?.size === 0) {
      editorEventSubscribers.delete(frameEditorId);
    }
  };
}
function canAccessIframeWindow(iframe) {
  if (!iframe) {
    return false;
  }
  try {
    void iframe.contentWindow?.location.href;
    return true;
  } catch {
    return false;
  }
}
function watchCrossOriginIframe(frameEditorId, getIframe, server, createIo) {
  let registered = false;
  const tryRegister = () => {
    const iframe = getIframe();
    if (!iframe?.src || canAccessIframeWindow(iframe)) {
      return false;
    }
    if (registered) {
      registerCrossOriginBridge(frameEditorId, iframe, server, createIo);
      return true;
    }
    registered = true;
    registerCrossOriginBridge(frameEditorId, iframe, server, createIo);
    return true;
  };
  tryRegister();
  const timer = window.setInterval(() => {
    if (tryRegister()) {
      window.clearInterval(timer);
    }
  }, 10);
  const observer = new MutationObserver(() => {
    tryRegister();
  });
  observer.observe(document.body, { childList: true, subtree: true });
  return () => {
    window.clearInterval(timer);
    observer.disconnect();
    if (registered) {
      unregisterCrossOriginBridge(frameEditorId);
    }
  };
}

// vendor/onlyoffice-web-comp/internal/editor/empty.ts
var emptyDocx = "DOCY;v5;7372;CQCAAgAACYYCAAAPewYAABDZBgAABwUHAAAGYBoAAAT7GgAABf8aAAAUyBwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBQAAAPEDAAAAJAAAAAABAAEBAQIBAgMBAwQBBAUBBQYBCgcBCwgBCAkBCQoBBgsBBwkEAAAApAEAAAJ8AAAAAAMAAAAAAQEBAwAAAAABAgIDAAAAAAEBAwMAAAAAAQEFAwAAAAABAAcGAAAAAgQAAAAACB4AAAAABhgAAABDAGEAbQBiAHIAaQBhACAATQBhAHQAaAAJAwAAAAABAQwGAAAAAgQAAAAADQMAAAAAAQAOBgAAAAIEoAUAAAgrAwAABAQAAABQBwAABQQAAAAEAAIABgQAAAAAAAAAAIUAAAABIgAAAGMAbwBtAHAAYQB0AGkAYgBpAGwAaQB0AHkATQBvAGQAZQACUAAAAGgAdAB0AHAAOgAvAC8AcwBjAGgAZQBtAGEAcwAuAG0AaQBjAHIAbwBzAG8AZgB0AC4AYwBvAG0ALwBvAGYAZgBpAGMAZQAvAHcAbwByAGQAAwQAAAAxADUAALUAAAABVAAAAG8AdgBlAHIAcgBpAGQAZQBUAGEAYgBsAGUAUwB0AHkAbABlAEYAbwBuAHQAUwBpAHoAZQBBAG4AZABKAHUAcwB0AGkAZgBpAGMAYQB0AGkAbwBuAAJQAAAAaAB0AHQAcAA6AC8ALwBzAGMAaABlAG0AYQBzAC4AbQBpAGMAcgBvAHMAbwBmAHQALgBjAG8AbQAvAG8AZgBmAGkAYwBlAC8AdwBvAHIAZAADAgAAADEAAI0AAAABLAAAAGUAbgBhAGIAbABlAE8AcABlAG4AVAB5AHAAZQBGAGUAYQB0AHUAcgBlAHMAAlAAAABoAHQAdABwADoALwAvAHMAYwBoAGUAbQBhAHMALgBtAGkAYwByAG8AcwBvAGYAdAAuAGMAbwBtAC8AbwBmAGYAaQBjAGUALwB3AG8AcgBkAAMCAAAAMQAAjQAAAAEsAAAAZABvAE4AbwB0AEYAbABpAHAATQBpAHIAcgBvAHIASQBuAGQAZQBuAHQAcwACUAAAAGgAdAB0AHAAOgAvAC8AcwBjAGgAZQBtAGEAcwAuAG0AaQBjAHIAbwBzAG8AZgB0AC4AYwBvAG0ALwBvAGYAZgBpAGMAZQAvAHcAbwByAGQAAwIAAAAxAACjAAAAAUIAAABkAGkAZgBmAGUAcgBlAG4AdABpAGEAdABlAE0AdQBsAHQAaQByAG8AdwBUAGEAYgBsAGUASABlAGEAZABlAHIAcwACUAAAAGgAdAB0AHAAOgAvAC8AcwBjAGgAZQBtAGEAcwAuAG0AaQBjAHIAbwBzAG8AZgB0AC4AYwBvAG0ALwBvAGYAZgBpAGMAZQAvAHcAbwByAGQAAwIAAAAxAAoCAAAALgALAgAAACwAAVkAAAD6AAsAAABOAG8AcgBtAGEAbAAuAGQAbwB0AG0ABAAAAAAFAQAAAAYAAAAABwAAAAAMAA0ADgAPAPsAGwAAAPoQAAAAABEAAAAAEgAAAAAUAAAAABYBAAAA+wInAAAA+gAAAAAABAAAAAAFAAAAAPsAEQAAAPoIAAAAAAoAAAAADQAAAAD7VxMAAANSEwAAFE0TAAD6AAsAAABPAGYAZgBpAGMAZQAgADtOmJgLIAsg+wAiEwAAABUBAAD6AAYAAABPAGYAZgBpAGMAZQD7AA0AAAABCAAAAPoAWwGbAtX7AQ0AAAABCAAAAPoA7QF9AjH7Ag0AAAABCAAAAPoApQGlAqX7Aw0AAAABCAAAAPoA/wHAAgD7BA0AAAABCAAAAPoARAFyAsT7BQ0AAAABCAAAAPoAcAGtAkf7CCYAAAAEIQAAAPoACgAAAHcAaQBuAGQAbwB3AFQAZQB4AHQAAQACAAMA+wkNAAAAAQgAAAD6AEQBVAJq+woNAAAAAQgAAAD6AJUBTwJy+wsNAAAAAQgAAAD6AAUBYwLB+wweAAAABBkAAAD6AAYAAAB3AGkAbgBkAG8AdwAB/wL/A//7DQ0AAAABCAAAAPoA5wHmAub7ATkLAAD6AAYAAABPAGYAZgBpAGMAZQD7AMIFAAAARAAAAPoBFAAAADAAMgAwAEYAMAAzADAAMgAwADIAMAAyADAANAAwADMAMAAyADAANAADCAAAAEl7v34gAEwAaQBnAGgAdAD7AREAAAD6AwUAAABBAHIAaQBhAGwA+wIRAAAA+gMFAAAAQQByAGkAYQBsAPsDSAUAAB4AAAAAKgAAAPoABAAAAEoAcABhAG4AAQsAAAA4brQwtzDDMK8wIABMAGkAZwBoAHQA+wAeAAAA+gAEAAAASABhAG4AZwABBQAAANG5QMcgAOCsFbX7ACQAAAD6AAQAAABIAGEAbgBzAAEIAAAASXu/fiAATABpAGcAaAB0APsAHAAAAPoABAAAAEgAYQBuAHQAAQQAAACwZTB9DmbUmvsAMgAAAPoABAAAAEEAcgBhAGIAAQ8AAABUAGkAbQBlAHMAIABOAGUAdwAgAFIAbwBtAGEAbgD7ADIAAAD6AAQAAABIAGUAYgByAAEPAAAAVABpAG0AZQBzACAATgBlAHcAIABSAG8AbQBhAG4A+wAqAAAA+gAEAAAAVABoAGEAaQABCwAAAEEAbgBnAHMAYQBuAGEAIABOAGUAdwD7AB4AAAD6AAQAAABFAHQAaABpAAEFAAAATgB5AGEAbABhAPsAIAAAAPoABAAAAEIAZQBuAGcAAQYAAABWAHIAaQBuAGQAYQD7ACAAAAD6AAQAAABHAHUAagByAAEGAAAAUwBoAHIAdQB0AGkA+wAmAAAA+gAEAAAASwBoAG0AcgABCQAAAE0AbwBvAGwAQgBvAHIAYQBuAPsAHgAAAPoABAAAAEsAbgBkAGEAAQUAAABUAHUAbgBnAGEA+wAeAAAA+gAEAAAARwB1AHIAdQABBQAAAFIAYQBhAHYAaQD7ACQAAAD6AAQAAABDAGEAbgBzAAEIAAAARQB1AHAAaABlAG0AaQBhAPsAPAAAAPoABAAAAEMAaABlAHIAARQAAABQAGwAYQBuAHQAYQBnAGUAbgBlAHQAIABDAGgAZQByAG8AawBlAGUA+wA4AAAA+gAEAAAAWQBpAGkAaQABEgAAAE0AaQBjAHIAbwBzAG8AZgB0ACAAWQBpACAAQgBhAGkAdABpAPsAOAAAAPoABAAAAFQAaQBiAHQAARIAAABNAGkAYwByAG8AcwBvAGYAdAAgAEgAaQBtAGEAbABhAHkAYQD7ACIAAAD6AAQAAABUAGgAYQBhAAEHAAAATQBWACAAQgBvAGwAaQD7ACAAAAD6AAQAAABEAGUAdgBhAAEGAAAATQBhAG4AZwBhAGwA+wAiAAAA+gAEAAAAVABlAGwAdQABBwAAAEcAYQB1AHQAYQBtAGkA+wAeAAAA+gAEAAAAVABhAG0AbAABBQAAAEwAYQB0AGgAYQD7ADYAAAD6AAQAAABTAHkAcgBjAAERAAAARQBzAHQAcgBhAG4AZwBlAGwAbwAgAEUAZABlAHMAcwBhAPsAIgAAAPoABAAAAE8AcgB5AGEAAQcAAABLAGEAbABpAG4AZwBhAPsAIgAAAPoABAAAAE0AbAB5AG0AAQcAAABLAGEAcgB0AGkAawBhAPsAJgAAAPoABAAAAEwAYQBvAG8AAQkAAABEAG8AawBDAGgAYQBtAHAAYQD7ACwAAAD6AAQAAABTAGkAbgBoAAEMAAAASQBzAGsAbwBvAGwAYQAgAFAAbwB0AGEA+wAyAAAA+gAEAAAATQBvAG4AZwABDwAAAE0AbwBuAGcAbwBsAGkAYQBuACAAQgBhAGkAdABpAPsAMgAAAPoABAAAAFYAaQBlAHQAAQ8AAABUAGkAbQBlAHMAIABOAGUAdwAgAFIAbwBtAGEAbgD7ADQAAAD6AAQAAABVAGkAZwBoAAEQAAAATQBpAGMAcgBvAHMAbwBmAHQAIABVAGkAZwBoAHUAcgD7ACIAAAD6AAQAAABHAGUAbwByAAEHAAAAUwB5AGwAZgBhAGUAbgD7AVoFAAAAOAAAAPoBFAAAADAAMgAwAEYAMAA1ADAAMgAwADIAMAAyADAANAAwADMAMAAyADAANAADAgAAAEl7v377AREAAAD6AwUAAABBAHIAaQBhAGwA+wIRAAAA+gMFAAAAQQByAGkAYQBsAPsD7AQAAB4AAAAAGgAAAPoABAAAAEoAcABhAG4AAQMAAAA4bg5mHWf7AB4AAAD6AAQAAABIAGEAbgBnAAEFAAAA0blAxyAA4KwVtfsAGAAAAPoABAAAAEgAYQBuAHMAAQIAAABJe79++wAcAAAA+gAEAAAASABhAG4AdAABBAAAALBlMH0OZtSa+wAeAAAA+gAEAAAAQQByAGEAYgABBQAAAEEAcgBpAGEAbAD7AB4AAAD6AAQAAABIAGUAYgByAAEFAAAAQQByAGkAYQBsAPsAKAAAAPoABAAAAFQAaABhAGkAAQoAAABDAG8AcgBkAGkAYQAgAE4AZQB3APsAHgAAAPoABAAAAEUAdABoAGkAAQUAAABOAHkAYQBsAGEA+wAgAAAA+gAEAAAAQgBlAG4AZwABBgAAAFYAcgBpAG4AZABhAPsAIAAAAPoABAAAAEcAdQBqAHIAAQYAAABTAGgAcgB1AHQAaQD7ACQAAAD6AAQAAABLAGgAbQByAAEIAAAARABhAHUAbgBQAGUAbgBoAPsAHgAAAPoABAAAAEsAbgBkAGEAAQUAAABUAHUAbgBnAGEA+wAeAAAA+gAEAAAARwB1AHIAdQABBQAAAFIAYQBhAHYAaQD7ACQAAAD6AAQAAABDAGEAbgBzAAEIAAAARQB1AHAAaABlAG0AaQBhAPsAPAAAAPoABAAAAEMAaABlAHIAARQAAABQAGwAYQBuAHQAYQBnAGUAbgBlAHQAIABDAGgAZQByAG8AawBlAGUA+wA4AAAA+gAEAAAAWQBpAGkAaQABEgAAAE0AaQBjAHIAbwBzAG8AZgB0ACAAWQBpACAAQgBhAGkAdABpAPsAOAAAAPoABAAAAFQAaQBiAHQAARIAAABNAGkAYwByAG8AcwBvAGYAdAAgAEgAaQBtAGEAbABhAHkAYQD7ACIAAAD6AAQAAABUAGgAYQBhAAEHAAAATQBWACAAQgBvAGwAaQD7ACAAAAD6AAQAAABEAGUAdgBhAAEGAAAATQBhAG4AZwBhAGwA+wAiAAAA+gAEAAAAVABlAGwAdQABBwAAAEcAYQB1AHQAYQBtAGkA+wAeAAAA+gAEAAAAVABhAG0AbAABBQAAAEwAYQB0AGgAYQD7ADYAAAD6AAQAAABTAHkAcgBjAAERAAAARQBzAHQAcgBhAG4AZwBlAGwAbwAgAEUAZABlAHMAcwBhAPsAIgAAAPoABAAAAE8AcgB5AGEAAQcAAABLAGEAbABpAG4AZwBhAPsAIgAAAPoABAAAAE0AbAB5AG0AAQcAAABLAGEAcgB0AGkAawBhAPsAJgAAAPoABAAAAEwAYQBvAG8AAQkAAABEAG8AawBDAGgAYQBtAHAAYQD7ACwAAAD6AAQAAABTAGkAbgBoAAEMAAAASQBzAGsAbwBvAGwAYQAgAFAAbwB0AGEA+wAyAAAA+gAEAAAATQBvAG4AZwABDwAAAE0AbwBuAGcAbwBsAGkAYQBuACAAQgBhAGkAdABpAPsAHgAAAPoABAAAAFYAaQBlAHQAAQUAAABBAHIAaQBhAGwA+wA0AAAA+gAEAAAAVQBpAGcAaAABEAAAAE0AaQBjAHIAbwBzAG8AZgB0ACAAVQBpAGcAaAB1AHIA+wAiAAAA+gAEAAAARwBlAG8AcgABBwAAAFMAeQBsAGYAYQBlAG4A+wLFBgAA+gAGAAAATwBmAGYAaQBjAGUA+wB2AwAAAwAAAAATAAAAAw4AAAAACQAAAAMEAAAA+gAO+wCmAQAABKEBAAD6AQH7AIoBAAADAAAAAH0AAAD6AAAAAAD7AHEAAAADbAAAAPoADvsAYwAAAAMAAAABHAAAAPoACAAAAGEAOgBsAHUAbQBNAG8AZAABsK0BAPsBHAAAAPoACAAAAGEAOgBzAGEAdABNAG8AZAABKJoBAPsBGAAAAPoABgAAAGEAOgB0AGkAbgB0AAG4BQEA+wB9AAAA+gBQwwAA+wBxAAAAA2wAAAD6AA77AGMAAAADAAAAARwAAAD6AAgAAABhADoAbAB1AG0ATQBvAGQAASiaAQD7ARwAAAD6AAgAAABhADoAcwBhAHQATQBvAGQAAViSAQD7ARgAAAD6AAYAAABhADoAdABpAG4AdAABKB0BAPsAfQAAAPoAoIYBAPsAcQAAAANsAAAA+gAO+wBjAAAAAwAAAAEcAAAA+gAIAAAAYQA6AGwAdQBtAE0AbwBkAAEomgEA+wEcAAAA+gAIAAAAYQA6AHMAYQB0AE0AbwBkAAHIqQEA+wEYAAAA+gAGAAAAYQA6AHQAaQBuAHQAAWg8AQD7AQkAAAD6AMBlUgABAPsAqgEAAASlAQAA+gEB+wCOAQAAAwAAAAB9AAAA+gAAAAAA+wBxAAAAA2wAAAD6AA77AGMAAAADAAAAARwAAAD6AAgAAABhADoAcwBhAHQATQBvAGQAAViSAQD7ARwAAAD6AAgAAABhADoAbAB1AG0ATQBvAGQAAXCOAQD7ARgAAAD6AAYAAABhADoAdABpAG4AdAABMG8BAPsAfwAAAPoAUMMAAPsAcwAAAANuAAAA+gAO+wBlAAAAAwAAAAEcAAAA+gAIAAAAYQA6AHMAYQB0AE0AbwBkAAGwrQEA+wEcAAAA+gAIAAAAYQA6AGwAdQBtAE0AbwBkAAGghgEA+wEaAAAA+gAHAAAAYQA6AHMAaABhAGQAZQABoIYBAPsAfwAAAPoAoIYBAPsAcwAAAANuAAAA+gAO+wBlAAAAAwAAAAEcAAAA+gAIAAAAYQA6AGwAdQBtAE0AbwBkAAG4ggEA+wEcAAAA+gAIAAAAYQA6AHMAYQB0AE0AbwBkAAHA1AEA+wEaAAAA+gAHAAAAYQA6AHMAaABhAGQAZQABsDABAPsBCQAAAPoAwGVSAAEA+wHQAAAAAwAAAAA/AAAA+gAAAQACAQPOGAAA+wATAAAAAw4AAAAACQAAAAMEAAAA+gAO+wEEAAAA+gAG+wIMAAAA+gADAAAAAQA1DAD7AD8AAAD6AAABAAIBA5wxAAD7ABMAAAADDgAAAAAJAAAAAwQAAAD6AA77AQQAAAD6AAb7AgwAAAD6AAMAAAABADUMAPsAPwAAAPoAAAEAAgEDakoAAPsAEwAAAAMOAAAAAAkAAAADBAAAAPoADvsBBAAAAPoABvsCDAAAAPoAAwAAAAEANQwA+wITAAAAAwAAAAAAAAAAAAAAAAAAAAAAAANFAgAAAwAAAAATAAAAAw4AAAAACQAAAAMEAAAA+gAO+wBaAAAAA1UAAAAAUAAAAANLAAAA+gAO+wBCAAAAAgAAAAEYAAAA+gAGAAAAYQA6AHQAaQBuAHQAARhzAQD7ARwAAAD6AAgAAABhADoAcwBhAHQATQBvAGQAARCYAgD7AMUBAAAEwAEAAPoBAfsAqQEAAAMAAAAAnAAAAPoAAAAAAPsAkAAAAAOLAAAA+gAO+wCCAAAABAAAAAEYAAAA+gAGAAAAYQA6AHQAaQBuAHQAAUhrAQD7ARwAAAD6AAgAAABhADoAcwBhAHQATQBvAGQAAfBJAgD7ARoAAAD6AAcAAABhADoAcwBoAGEAZABlAAHQfgEA+wEcAAAA+gAIAAAAYQA6AGwAdQBtAE0AbwBkAAFwjgEA+wCcAAAA+gBQwwAA+wCQAAAAA4sAAAD6AA77AIIAAAAEAAAAARgAAAD6AAYAAABhADoAdABpAG4AdAAB0H4BAPsBHAAAAPoACAAAAGEAOgBzAGEAdABNAG8AZAAB0PsBAPsBGgAAAPoABwAAAGEAOgBzAGgAYQBkAGUAAZBfAQD7ARwAAAD6AAgAAABhADoAbAB1AG0ATQBvAGQAAViSAQD7AF4AAAD6AKCGAQD7AFIAAAADTQAAAPoADvsARAAAAAIAAAABGgAAAPoABwAAAGEAOgBzAGgAYQBkAGUAARj2AAD7ARwAAAD6AAgAAABhADoAcwBhAHQATQBvAGQAAcDUAQD7AQkAAAD6AMBlUgABAPsEBAAAAAAAAACXAAAAADQAAAACLwAAABccAAAAAAQAAAAAAAAAAQ4AAABfAEcAbwBCAGEAYwBrABgJAAAAAAQAAAAAAAAABFkAAAAADAAAAAMEgi4AAAQExkEAAAEqAAAABgQIBwAABwSgBQAACAQIBwAACQSgBQAACgRTAwAACwTgAwAADAQAAAAAAgYAAAAAAQABAQAICQAAAAMEAAAAqQEAAAAAAADFAQAAAUgAAAAoAQcpAQcqAQYrAQUIBBUAAAAWBBYAAAAZBgoAAABlAG4ALQBVAFMAGgYKAAAAYQByAC0AUwBBABsGCgAAAHoAaAAtAEMATgACcwEAAAA1AAAAAQIAAABhAAIMAAAATgBvAHIAbQBhAGwACQEAAAADCAEAAAABCgEAAAABBgYAAAAZAQAFAQMAWwAAAAEEAAAAYQAwAAIsAAAARABlAGYAYQB1AGwAdAAgAFAAYQByAGEAZwByAGEAcABoACAARgBvAG4AdAAJAQAAAAEIAQAAAAELBAAAAAEAAAANAQAAAAEOAQAAAAEAkgAAAAEEAAAAYQAxAAIYAAAATgBvAHIAbQBhAGwAIABUAGEAYgBsAGUACQEAAAAECAEAAAABCwQAAABjAAAADQEAAAABDgEAAAABB0YAAAATBAAAAAAAAAAFOAAAAAAJAAAAAAEBAgRsAAAAAQkAAAAAAQECBAAAAAACCQAAAAABAQIEbAAAAAMJAAAAAAEBAgQAAAAAAD0AAAABBAAAAGEAMgACDgAAAE4AbwAgAEwAaQBzAHQACQEAAAACCAEAAAABCwQAAABjAAAADQEAAAABDgEAAAABAAAAAA==";
var emptyXlsx = "XLSY;v2;6160;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";
var emptyPptx = "PPTY;v1;47829;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";
var emptyPdf = '%PDF-1.7\n%âãÏÓ\n5 0 obj\n<<\n/Filter /FlateDecode\n/Length 8\n>>\nstream\nx\0\0\0\0\nendstream\nendobj\n1 0 obj\n<<\n/Type /Catalog\n/Pages 2 0 R\n>>\nendobj\n7 0 obj\n<<\n/Type /ObjStm\n/N 2\n/First 9\n/Filter /FlateDecode\n/Length 113\n>>\nstream\nx3R0P0V01â²±áÒ©,HUÐHLO-æÒ÷ÎL)V6Ê\x07Åré;çæ(rÙÙUúæ§¸$¤*h¸XXZDiré\x07å§&§)h¤¤äçç+»xkô\0EÔ\nendstream\nendobj\n6 0 obj\n<<\n/Type /ObjStm\n/N 1\n/First 4\n/Filter /FlateDecode\n/Length 98\n>>\nstream\nx3Q0à²±áÒ©,HUÐHLOåÒw.Ê/pÊ¯P6P0P034R0·4åÒ÷MMÉLÄ"Z_ZZ¬\x004ÈÎ¨??¯$5¯¤XÁ¨.\vhj«`æ\0\0ø\nendstream\nendobj\n8 0 obj\n<<\n/Size 9\n/Root 1 0 R\n/Info 3 0 R\n/ID [<A0D2A09C7D58A60922A6A67A45FCEFD4> <BC61A8F4C538C967CA2B1ED93A05AF53>]\n/Type /XRef\n/W [1 2 2]\n/Filter /FlateDecode\n/Index [0 9]\n/Length 44\n>>\nstream\nxc``øÿ!B021°1002ð	Æd«H0i00\0\0o§\nendstream\nendobj\nstartxref\n552\n%%EOF\n';

// vendor/onlyoffice-web-comp/internal/editor/office-format.ts
var crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 3988292384 ^ c >>> 1 : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();
function crc32(data) {
  let crc = 4294967295;
  for (const byte of data) {
    crc = crcTable[(crc ^ byte) & 255] ^ crc >>> 8;
  }
  return (crc ^ 4294967295) >>> 0;
}
function readU16(data, offset) {
  return data[offset] | data[offset + 1] << 8;
}
function readU32(data, offset) {
  return (data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24) >>> 0;
}
function writeU16(data, offset, value) {
  data[offset] = value & 255;
  data[offset + 1] = value >>> 8 & 255;
}
function writeU32(data, offset, value) {
  data[offset] = value & 255;
  data[offset + 1] = value >>> 8 & 255;
  data[offset + 2] = value >>> 16 & 255;
  data[offset + 3] = value >>> 24 & 255;
}
function concatBytes(parts) {
  const size = parts.reduce((total, part) => total + part.length, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}
async function decompressDeflateRaw(data) {
  if (!("DecompressionStream" in globalThis)) {
    return null;
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
function readZipEntries(data) {
  const input = new Uint8Array(data);
  let eocd = -1;
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65558); i--) {
    if (readU32(input, i) === 101010256) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;
  const entryCount = readU16(input, eocd + 10);
  const centralOffset = readU32(input, eocd + 16);
  const decoder = new TextDecoder();
  const entries2 = [];
  let readOffset = centralOffset;
  for (let i = 0; i < entryCount; i++) {
    if (readU32(input, readOffset) !== 33639248) return null;
    const method = readU16(input, readOffset + 10);
    const modTime = readU16(input, readOffset + 12);
    const modDate = readU16(input, readOffset + 14);
    const crc = readU32(input, readOffset + 16);
    const compressedSize = readU32(input, readOffset + 20);
    const uncompressedSize = readU32(input, readOffset + 24);
    const nameLength = readU16(input, readOffset + 28);
    const extraLength = readU16(input, readOffset + 30);
    const commentLength = readU16(input, readOffset + 32);
    const internalAttrs = readU16(input, readOffset + 36);
    const externalAttrs = readU32(input, readOffset + 38);
    const localOffset = readU32(input, readOffset + 42);
    const nameBytes = input.slice(readOffset + 46, readOffset + 46 + nameLength);
    const name = decoder.decode(nameBytes);
    if (readU32(input, localOffset) !== 67324752) return null;
    const localNameLength = readU16(input, localOffset + 26);
    const localExtraLength = readU16(input, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressedData = input.slice(dataStart, dataStart + compressedSize);
    entries2.push({
      name,
      nameBytes,
      method,
      modTime,
      modDate,
      crc,
      compressedSize,
      uncompressedSize,
      compressedData,
      internalAttrs,
      externalAttrs
    });
    readOffset += 46 + nameLength + extraLength + commentLength;
  }
  return entries2;
}
async function readZipEntryData(entry) {
  if (entry.method === 0) {
    return entry.compressedData;
  }
  if (entry.method === 8) {
    return decompressDeflateRaw(entry.compressedData);
  }
  return null;
}
function buildLocalHeader(entry, crc, compressedSize, uncompressedSize) {
  const localHeader = new Uint8Array(30 + entry.nameBytes.length);
  writeU32(localHeader, 0, 67324752);
  writeU16(localHeader, 4, 20);
  writeU16(localHeader, 6, 0);
  writeU16(localHeader, 8, entry.method);
  writeU16(localHeader, 10, entry.modTime);
  writeU16(localHeader, 12, entry.modDate);
  writeU32(localHeader, 14, crc);
  writeU32(localHeader, 18, compressedSize);
  writeU32(localHeader, 22, uncompressedSize);
  writeU16(localHeader, 26, entry.nameBytes.length);
  writeU16(localHeader, 28, 0);
  localHeader.set(entry.nameBytes, 30);
  return localHeader;
}
function buildCentralHeader(entry, crc, compressedSize, uncompressedSize, localOffset) {
  const centralHeader = new Uint8Array(46 + entry.nameBytes.length);
  writeU32(centralHeader, 0, 33639248);
  writeU16(centralHeader, 4, 20);
  writeU16(centralHeader, 6, 20);
  writeU16(centralHeader, 8, 0);
  writeU16(centralHeader, 10, entry.method);
  writeU16(centralHeader, 12, entry.modTime);
  writeU16(centralHeader, 14, entry.modDate);
  writeU32(centralHeader, 16, crc);
  writeU32(centralHeader, 20, compressedSize);
  writeU32(centralHeader, 24, uncompressedSize);
  writeU16(centralHeader, 28, entry.nameBytes.length);
  writeU16(centralHeader, 30, 0);
  writeU16(centralHeader, 32, 0);
  writeU16(centralHeader, 34, 0);
  writeU16(centralHeader, 36, entry.internalAttrs);
  writeU32(centralHeader, 38, entry.externalAttrs);
  writeU32(centralHeader, 42, localOffset);
  centralHeader.set(entry.nameBytes, 46);
  return centralHeader;
}
function isOfficeXmlEntryName(name) {
  return /\.(xml|rels)$/i.test(name);
}
function getZipXmlUncompressedSize(entries2) {
  let totalSize = 0;
  let entryCount = 0;
  for (const entry of entries2) {
    if (entry.name.endsWith("/") || !isOfficeXmlEntryName(entry.name)) {
      continue;
    }
    totalSize += entry.uncompressedSize;
    entryCount += 1;
  }
  return { totalSize, entryCount };
}
function writeZipEntries(entries2, replacements) {
  const encoder2 = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let outputOffset = 0;
  const emitted = /* @__PURE__ */ new Set();
  let emittedCount = 0;
  const emit = (entry, replacement) => {
    const outputData = replacement?.data ?? entry.compressedData;
    const method = replacement ? 0 : entry.method;
    const crc = replacement ? crc32(outputData) : entry.crc;
    const uncompressedSize = replacement ? outputData.length : entry.uncompressedSize;
    const modTime = replacement?.modTime ?? entry.modTime;
    const modDate = replacement?.modDate ?? entry.modDate;
    const nextEntry = { ...entry, method, modTime, modDate };
    const localHeader = buildLocalHeader(
      nextEntry,
      crc,
      outputData.length,
      uncompressedSize
    );
    localParts.push(localHeader, outputData);
    centralParts.push(
      buildCentralHeader(
        nextEntry,
        crc,
        outputData.length,
        uncompressedSize,
        outputOffset
      )
    );
    outputOffset += localHeader.length + outputData.length;
    emitted.add(entry.name);
    emittedCount += 1;
  };
  for (const entry of entries2) {
    emit(entry, replacements.get(entry.name));
  }
  for (const [name, replacement] of replacements) {
    if (emitted.has(name)) continue;
    const nameBytes = encoder2.encode(name);
    emit({
      name,
      nameBytes,
      method: 0,
      modTime: replacement.modTime ?? 0,
      modDate: replacement.modDate ?? 0,
      crc: 0,
      compressedData: replacement.data,
      uncompressedSize: replacement.data.length,
      internalAttrs: 0,
      externalAttrs: 0
    });
  }
  const centralDirectory = concatBytes(centralParts);
  const end = new Uint8Array(22);
  writeU32(end, 0, 101010256);
  writeU16(end, 8, emittedCount);
  writeU16(end, 10, emittedCount);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, outputOffset);
  return concatBytes([...localParts, centralDirectory, end]).buffer;
}
async function convertCsvBufferToXlsxBuffer(buffer) {
  return createXlsxFromRows(parseCsvBuffer(buffer));
}
var encoder = new TextEncoder();
function utf8(value) {
  return encoder.encode(value);
}
function xml(strings, ...values) {
  return strings.reduce((result, value, index) => `${result}${value}${values[index] ?? ""}`, "").trim();
}
function escapeXml(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function decodeText(buffer) {
  return new TextDecoder("utf-8").decode(buffer);
}
function zip(entries2) {
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;
  for (const [name, content] of entries2) {
    const nameBytes = utf8(name);
    const data = typeof content === "string" ? utf8(content) : content;
    const crc = crc32(data);
    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeU32(localHeader, 0, 67324752);
    writeU16(localHeader, 4, 20);
    writeU16(localHeader, 6, 0);
    writeU16(localHeader, 8, 0);
    writeU16(localHeader, 10, 0);
    writeU16(localHeader, 12, 0);
    writeU32(localHeader, 14, crc);
    writeU32(localHeader, 18, data.length);
    writeU32(localHeader, 22, data.length);
    writeU16(localHeader, 26, nameBytes.length);
    writeU16(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, data);
    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeU32(centralHeader, 0, 33639248);
    writeU16(centralHeader, 4, 20);
    writeU16(centralHeader, 6, 20);
    writeU16(centralHeader, 8, 0);
    writeU16(centralHeader, 10, 0);
    writeU16(centralHeader, 12, 0);
    writeU16(centralHeader, 14, 0);
    writeU32(centralHeader, 16, crc);
    writeU32(centralHeader, 20, data.length);
    writeU32(centralHeader, 24, data.length);
    writeU16(centralHeader, 28, nameBytes.length);
    writeU16(centralHeader, 30, 0);
    writeU16(centralHeader, 32, 0);
    writeU16(centralHeader, 34, 0);
    writeU16(centralHeader, 36, 0);
    writeU32(centralHeader, 38, 0);
    writeU32(centralHeader, 42, localOffset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);
    localOffset += localHeader.length + data.length;
  }
  const centralDirectory = concatBytes(centralParts);
  const end = new Uint8Array(22);
  writeU32(end, 0, 101010256);
  writeU16(end, 8, entries2.length);
  writeU16(end, 10, entries2.length);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, localOffset);
  return concatBytes([...localParts, centralDirectory, end]).buffer;
}
var coreProps = xml`
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>Plain text fallback</dc:title>
  </cp:coreProperties>
`;
var appProps = xml`
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>OnlyOffice Web Comp</Application>
  </Properties>
`;
function createDocxFromText(buffer) {
  const paragraphs = decodeText(buffer).split(/\r\n|\r|\n/).map((line) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`).join("");
  return zip([
    [
      "[Content_Types].xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
      `
    ],
    [
      "_rels/.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
      `
    ],
    ["docProps/core.xml", coreProps],
    ["docProps/app.xml", appProps],
    [
      "word/document.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            ${paragraphs || "<w:p><w:r><w:t></w:t></w:r></w:p>"}
            <w:sectPr>
              <w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>
          </w:body>
        </w:document>
      `
    ],
    [
      "word/_rels/document.xml.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
      `
    ]
  ]);
}
function createXlsxFromRows(rows) {
  const sheetData = rows.map((fields, rowIndex) => {
    const rowNumber = rowIndex + 1;
    if (fields.length === 0) {
      return `<row r="${rowNumber}"/>`;
    }
    const cells = fields.map((value, columnIndex) => {
      const ref = `${xlsxColumnName(columnIndex)}${rowNumber}`;
      const text = escapeXml(String(value ?? "").slice(0, 32767));
      return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`;
    }).join("");
    return `<row r="${rowNumber}">${cells}</row>`;
  }).join("");
  return zip([
    [
      "[Content_Types].xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        </Types>
      `
    ],
    [
      "_rels/.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
      `
    ],
    ["docProps/core.xml", coreProps],
    ["docProps/app.xml", appProps],
    [
      "xl/workbook.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
      `
    ],
    [
      "xl/_rels/workbook.xml.rels",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
      `
    ],
    [
      "xl/styles.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
          <borders count="1"><border/></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        </styleSheet>
      `
    ],
    [
      "xl/worksheets/sheet1.xml",
      xml`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>${sheetData}</sheetData>
        </worksheet>
      `
    ]
  ]);
}
function createXlsxFromText(buffer) {
  return createXlsxFromRows([[decodeText(buffer)]]);
}
function xlsxColumnName(index) {
  let name = "";
  let value = index;
  do {
    name = String.fromCharCode(65 + value % 26) + name;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return name;
}

// vendor/onlyoffice-web-comp/util/document-file.ts
function getOnlyOfficeMimeType(fileExtension2) {
  const extension = (fileExtension2 || "").toLowerCase();
  if (["xlsx", "xls"].includes(extension)) {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
  if (extension === "csv") return "text/csv";
  if (extension === "docm") return "application/vnd.ms-word.document.macroEnabled.12";
  if (["docx"].includes(extension)) {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document;charset=UTF-8";
  }
  if (["pptx", "ppt"].includes(extension)) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  if (extension === "pdf") {
    return "application/pdf";
  }
  if (extension === "txt") return "text/plain";
  return "application/octet-stream";
}

// vendor/onlyoffice-web-comp/internal/editor/server.ts
var PROGRAMMATIC_EXPORT_ACK_URL = "onlyoffice://export/ack";
function mergeBuffers(buffers) {
  const totalLength = buffers.reduce((acc, buffer) => acc + buffer.length, 0);
  const mergedBuffer = new Uint8Array(totalLength);
  let offset = 0;
  for (const buffer of buffers) {
    mergedBuffer.set(buffer, offset);
    offset += buffer.length;
  }
  return mergedBuffer;
}
function isValidEditorBin(data) {
  if (data.length < 4) {
    return false;
  }
  const magic = String.fromCharCode(data[0], data[1], data[2], data[3]);
  return magic === "XLSY" || magic === "DOCY" || magic === "PPTY";
}
function isPdfBytes(data) {
  return data.length >= 5 && data[0] === 37 && data[1] === 80 && data[2] === 68 && data[3] === 70 && data[4] === 45;
}
function isOfficeZipFileType(fileType) {
  return /^(docx|docm|dotx|dotm|xlsx|xlsm|xltx|xltm|pptx|pptm|potx|potm|ppsx|ppsm)$/i.test(
    getFileExt(fileType)
  );
}
function isZipBytes(data) {
  return data.length >= 4 && data[0] === 80 && data[1] === 75 && (data[2] === 3 && data[3] === 4 || data[2] === 5 && data[3] === 6 || data[2] === 7 && data[3] === 8);
}
function createPlainTextOfficeFallback(buffer, fileType) {
  switch (getDocumentType(fileType)) {
    case "word":
      return { buffer: createDocxFromText(buffer), fileType: "docx" };
    case "cell":
      return { buffer: createXlsxFromText(buffer), fileType: "xlsx" };
    default:
      return null;
  }
}
var PDF_EOF_MARKER = new TextEncoder().encode("%%EOF");
function indexOfSubarray(haystack, needle) {
  if (needle.length === 0) {
    return 0;
  }
  outer: for (let i = 0; i <= haystack.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (haystack[i + j] !== needle[j]) {
        continue outer;
      }
    }
    return i;
  }
  return -1;
}
function assertValidPdfOutput(data, label = "PDF") {
  if (!isPdfBytes(data)) {
    throw new Error(`${label} output is not a valid PDF`);
  }
  const scanLen = Math.min(data.byteLength, 1024 * 1024);
  const tail = data.subarray(data.byteLength - scanLen);
  if (indexOfSubarray(tail, PDF_EOF_MARKER) < 0) {
    throw new Error(`${label} output is corrupt (missing %%EOF trailer)`);
  }
}
function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      window.setTimeout(
        () => reject(new Error(`${label} timed out after ${ms}ms`)),
        ms
      );
    })
  ]);
}
var PDF_CONVERT_TIMEOUT_MS = 12e4;
var TRANSPARENT_PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
  ),
  (value) => value.charCodeAt(0)
);
function randomId() {
  return Math.random().toString(36).substring(2, 9);
}
function normalizeCoAuthoringLockBlocks(block) {
  if (Array.isArray(block)) {
    return block;
  }
  return [block];
}
function getCoAuthoringLockKey(block, isSpreadsheet) {
  if (isSpreadsheet && typeof block === "object" && block !== null && block.guid != null) {
    return String(block.guid);
  }
  return String(block);
}
function buildCoAuthoringLocks(block, fileType, userId) {
  const isSpreadsheet = getDocumentType(fileType) === "cell";
  const time = +/* @__PURE__ */ new Date();
  const locks = {};
  for (const item of normalizeCoAuthoringLockBlocks(block)) {
    const key = getCoAuthoringLockKey(item, isSpreadsheet);
    locks[key] = {
      time,
      user: userId,
      block: item
    };
  }
  return locks;
}
function getUrl(data, type) {
  const blob = new Blob([data], {
    type: type || "application/octet-stream"
  });
  return URL.createObjectURL(blob);
}
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function ensurePngContentType(contentTypes) {
  if (/Extension="png"/i.test(contentTypes)) {
    return contentTypes;
  }
  return contentTypes.replace(
    "</Types>",
    '<Default Extension="png" ContentType="image/png"/></Types>'
  );
}
function parseSvgSize(svgText) {
  const width = Number(/<svg\b[^>]*\bwidth="([\d.]+)/i.exec(svgText)?.[1]);
  const height = Number(/<svg\b[^>]*\bheight="([\d.]+)/i.exec(svgText)?.[1]);
  if (width > 0 && height > 0) {
    return { width, height };
  }
  const viewBox = /<svg\b[^>]*\bviewBox="([^"]+)"/i.exec(svgText)?.[1]?.trim().split(/\s+/).map(Number);
  if (viewBox?.length === 4 && viewBox[2] > 0 && viewBox[3] > 0) {
    return { width: viewBox[2], height: viewBox[3] };
  }
  return { width: 256, height: 256 };
}
async function rasterizeImageToPng(imageBytes, mimeType, fallbackSize) {
  if (typeof document === "undefined") {
    return TRANSPARENT_PNG;
  }
  const blob = new Blob([imageBytes], {
    type: mimeType
  });
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image();
    image.decoding = "async";
    const loaded = new Promise((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Failed to decode SVG image"));
    });
    image.src = url;
    await loaded;
    const width = image.naturalWidth || fallbackSize?.width || 256;
    const height = image.naturalHeight || fallbackSize?.height || 256;
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.ceil(width));
    canvas.height = Math.max(1, Math.ceil(height));
    const context = canvas.getContext("2d");
    if (!context) {
      return TRANSPARENT_PNG;
    }
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    const pngBlob = await new Promise(
      (resolve) => canvas.toBlob(resolve, "image/png")
    );
    if (!pngBlob) {
      return TRANSPARENT_PNG;
    }
    return new Uint8Array(await pngBlob.arrayBuffer());
  } catch (err) {
    console.warn("[EditorServer] SVG to PNG fallback:", err);
    return TRANSPARENT_PNG;
  } finally {
    URL.revokeObjectURL(url);
  }
}
async function rasterizeSvgToPng(svgBytes) {
  const svgText = new TextDecoder().decode(svgBytes);
  return rasterizeImageToPng(svgBytes, "image/svg+xml", parseSvgSize(svgText));
}
async function rasterizeWebpToPng(webpBytes) {
  return rasterizeImageToPng(webpBytes, "image/webp");
}
function replaceBytes(data, from, to) {
  if (from.byteLength !== to.byteLength) {
    throw new Error("Replacement must keep binary length unchanged");
  }
  const output = data.slice();
  for (let i = 0; i <= output.byteLength - from.byteLength; i++) {
    let matches = true;
    for (let j = 0; j < from.byteLength; j++) {
      if (output[i + j] !== from[j]) {
        matches = false;
        break;
      }
    }
    if (matches) {
      output.set(to, i);
      i += from.byteLength - 1;
    }
  }
  return output;
}
function replaceAsciiExtension(data, fromExt, toExt) {
  return replaceBytes(
    data,
    new TextEncoder().encode(fromExt),
    new TextEncoder().encode(toExt)
  );
}
function replaceUtf16LeExtension(data, fromExt, toExt) {
  const encode = (value) => {
    const bytes = new Uint8Array(value.length * 2);
    for (let i = 0; i < value.length; i++) {
      bytes[i * 2] = value.charCodeAt(i);
    }
    return bytes;
  };
  return replaceBytes(data, encode(fromExt), encode(toExt));
}
function rewriteEditorBinUnsupportedMediaRefs(data) {
  let output = data;
  for (const fromExt of [".emf", ".EMF", ".svg", ".SVG"]) {
    const toExt = fromExt === fromExt.toUpperCase() ? ".PNG" : ".png";
    output = replaceAsciiExtension(output, fromExt, toExt);
    output = replaceUtf16LeExtension(output, fromExt, toExt);
  }
  return output;
}
function buildContentDisposition(fileName) {
  const encoded = encodeURIComponent(fileName).replace(/['()]/g, escape);
  const ascii = fileName.replace(/[^\x20-\x7E]/g, "_").replace(/"/g, '\\"');
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encoded}`;
}
function getCacheDownloadName(outputName, title, filetype, downloadFileNames) {
  const stored = downloadFileNames.get(outputName);
  if (stored) {
    return stored;
  }
  if (outputName === "Editor.bin" || outputName.startsWith("output.")) {
    return ensureTitleWithExtension(title, filetype);
  }
  return outputName.split("/").pop() || outputName;
}
function getCacheResponseMimeType(name, fallbackFileType) {
  if (name.startsWith("media/")) {
    return detectImageMimeForName(name);
  }
  const filetype = getFileExt(name.replace(/^output\./, "")) || fallbackFileType;
  if (name === "Editor.bin") {
    return "application/octet-stream";
  }
  return getOnlyOfficeMimeType(filetype);
}
function getDataUrl(data, mimeType = "application/octet-stream") {
  let binary = "";
  const chunkSize = 32768;
  for (let i = 0; i < data.length; i += chunkSize) {
    binary += String.fromCharCode(...data.subarray(i, i + chunkSize));
  }
  return `data:${mimeType};base64,${btoa(binary)}`;
}
function parseClipboardImage(input) {
  const match = /^data:([^;]+);base64,(.+)$/.exec(input);
  if (!match) {
    return null;
  }
  const mime = match[1];
  const binary = atob(match[2]);
  const data = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    data[i] = binary.charCodeAt(i);
  }
  const subtype = mime.split("/")[1]?.split("+")[0] || "png";
  const ext = subtype === "jpeg" ? "jpg" : subtype;
  return { data, mime, ext };
}
async function fetchClipboardImage(input) {
  let url;
  try {
    url = new URL(input, window.location.href);
  } catch {
    return null;
  }
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    return null;
  }
  const response = await fetch(url.href);
  if (!response.ok) {
    return null;
  }
  const data = new Uint8Array(await response.arrayBuffer());
  const headerMime = response.headers.get("Content-Type")?.split(";")[0] ?? "";
  const mime = headerMime.startsWith("image/") ? headerMime : detectImageMime(data);
  if (!mime.startsWith("image/")) {
    return null;
  }
  const ext = getFileExt(url.pathname) || getImageExt(mime);
  return { data, mime, ext };
}
function detectImageMime(data) {
  if (data.length >= 8 && data[0] === 137 && data[1] === 80 && data[2] === 78 && data[3] === 71) {
    return "image/png";
  }
  if (data.length >= 3 && data[0] === 255 && data[1] === 216 && data[2] === 255) {
    return "image/jpeg";
  }
  if (data.length >= 6 && data[0] === 71 && data[1] === 73 && data[2] === 70) {
    return "image/gif";
  }
  if (data.length >= 12 && data[0] === 82 && data[1] === 73 && data[2] === 70 && data[3] === 70 && data[8] === 87 && data[9] === 69 && data[10] === 66 && data[11] === 80) {
    return "image/webp";
  }
  return "image/png";
}
function detectImageMimeForName(name) {
  switch (getFileExt(name)) {
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "webp":
      return "image/webp";
    case "png":
    default:
      return "image/png";
  }
}
function getImageExt(mime) {
  const subtype = mime.split("/")[1]?.split("+")[0] || "png";
  return subtype === "jpeg" ? "jpg" : subtype;
}
function asRecord(value) {
  return typeof value === "object" && value !== null ? value : null;
}
function getStringField(record, key) {
  const value = record?.[key];
  return typeof value === "string" ? value : void 0;
}
function getNumberField(record, key) {
  const value = record?.[key];
  return typeof value === "number" ? value : void 0;
}
function getBooleanField(record, key) {
  const value = record?.[key];
  return typeof value === "boolean" ? value : void 0;
}
function truncateText(value, maxLength = 80) {
  return value.length > maxLength ? `${value.slice(0, maxLength)}...` : value;
}
function tryParseJson(value) {
  const trimmed = value.trim();
  if (!trimmed || !/^[{[]/.test(trimmed)) {
    return value;
  }
  try {
    return JSON.parse(trimmed);
  } catch {
    return value;
  }
}
function getJsonArrayField(record, key) {
  const value = record?.[key];
  if (Array.isArray(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return void 0;
  }
  const parsed = tryParseJson(value);
  return Array.isArray(parsed) ? parsed : void 0;
}
function getPayloadByteLength(value) {
  if (value instanceof ArrayBuffer) {
    return value.byteLength;
  }
  if (ArrayBuffer.isView(value)) {
    return value.byteLength;
  }
  if (typeof value === "string") {
    return value.length;
  }
  return void 0;
}
function describeValue(value, maxLength = 120) {
  if (value == null) {
    return void 0;
  }
  if (typeof value === "string") {
    return truncateText(value, maxLength);
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (Array.isArray(value)) {
    return `array(${value.length})`;
  }
  if (typeof value === "object") {
    const record = value;
    const keys = Object.keys(record).slice(0, 6);
    return `{${keys.join(", ")}}`;
  }
  return typeof value;
}
function collectTextFragments(value, fragments = [], depth = 0) {
  if (fragments.length >= 3 || depth > 4 || value == null) {
    return fragments;
  }
  if (typeof value === "string") {
    const parsed = tryParseJson(value);
    if (parsed !== value) {
      return collectTextFragments(parsed, fragments, depth + 1);
    }
    const text = value.trim();
    if (text.length > 0 && text.length <= 200 && /[\p{L}\p{N}\u4e00-\u9fff]/u.test(text)) {
      fragments.push(truncateText(text));
    }
    return fragments;
  }
  if (Array.isArray(value)) {
    for (const item of value.slice(0, 8)) {
      collectTextFragments(item, fragments, depth + 1);
      if (fragments.length >= 3) break;
    }
    return fragments;
  }
  if (typeof value === "object") {
    const record = value;
    const preferredKeys = [
      "text",
      "Text",
      "value",
      "insert",
      "m_sText",
      "data",
      "items"
    ];
    const keys = [
      ...preferredKeys.filter((key) => key in record),
      ...Object.keys(record).filter((key) => !preferredKeys.includes(key))
    ];
    for (const key of keys.slice(0, 12)) {
      collectTextFragments(record[key], fragments, depth + 1);
      if (fragments.length >= 3) break;
    }
  }
  return fragments;
}
function summarizeLockBlock(block) {
  const blocks = normalizeCoAuthoringLockBlocks(block);
  return {
    lockCount: blocks.length,
    locks: blocks.slice(0, 3).map((item) => {
      if (typeof item === "object" && item !== null) {
        const record = item;
        return {
          guid: getStringField(record, "guid"),
          sheetId: record.sheetId ?? record.sheet ?? record.tabId,
          range: record.range ?? record.bbox ?? record.bounds,
          raw: describeValue(record)
        };
      }
      return { raw: describeValue(item) };
    })
  };
}
function getBase64ByteLength(value) {
  const normalized = value.replace(/\s/g, "");
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(normalized)) {
    return void 0;
  }
  const padding = normalized.endsWith("==") ? 2 : normalized.endsWith("=") ? 1 : 0;
  return Math.floor(normalized.length * 3 / 4) - padding;
}
function summarizeChangePacket(value) {
  if (typeof value !== "string") {
    return {
      kind: typeof value,
      byteLength: getPayloadByteLength(value),
      preview: describeValue(value)
    };
  }
  const separatorIndex = value.indexOf(";");
  if (separatorIndex < 0) {
    return {
      kind: "string",
      charLength: value.length,
      preview: truncateText(value)
    };
  }
  const declaredLength = Number(value.slice(0, separatorIndex));
  const base64 = value.slice(separatorIndex + 1);
  const byteLength = getBase64ByteLength(base64);
  return {
    kind: "binary-change",
    declaredLength: Number.isFinite(declaredLength) ? declaredLength : void 0,
    byteLength,
    byteLengthMatches: Number.isFinite(declaredLength) && byteLength != null ? declaredLength === byteLength : void 0,
    base64Preview: truncateText(base64, 32)
  };
}
function summarizeSaveChanges(record) {
  const changes = getJsonArrayField(record, "changes") ?? getJsonArrayField(record, "changesData") ?? getJsonArrayField(record, "data");
  const firstChange = changes?.[0] ?? record.change ?? record.data;
  const parsedFirstChange = typeof firstChange === "string" ? tryParseJson(firstChange) : firstChange;
  const textPreview = collectTextFragments(changes ?? firstChange);
  const packets = changes?.map(summarizeChangePacket);
  const changeBytes = packets?.reduce((total, packet) => total + (packet.byteLength ?? 0), 0) ?? getPayloadByteLength(firstChange);
  return {
    start: getBooleanField(record, "startSaveChanges"),
    end: getBooleanField(record, "endSaveChanges"),
    isCoAuthoring: getBooleanField(record, "isCoAuthoring"),
    isExcel: getBooleanField(record, "isExcel"),
    deleteIndex: getNumberField(record, "deleteIndex"),
    unlock: getBooleanField(record, "unlock"),
    releaseLocks: getBooleanField(record, "releaseLocks"),
    changesCount: changes?.length,
    changeBytes,
    packets: packets?.slice(0, 4),
    firstChange: describeValue(parsedFirstChange),
    textPreview
  };
}
function getSaveTypeLabel(value) {
  switch (value) {
    case 0 /* PartStart */:
      return "PartStart";
    case 1 /* Part */:
      return "Part";
    case 2 /* Complete */:
      return "Complete";
    case 3 /* CompleteAll */:
      return "CompleteAll";
    default:
      return value == null ? void 0 : String(value);
  }
}
function summarizeWsPayload(payloads) {
  const first = payloads[0];
  const record = asRecord(first);
  const data = asRecord(record?.data);
  const type = getStringField(record, "type");
  const dataType = getStringField(data, "type");
  const extraParts = [];
  let extra = {};
  if (type === "saveChanges" && record) {
    const saveSummary = summarizeSaveChanges(record);
    extra = saveSummary;
    if (saveSummary.changesCount != null) {
      extraParts.push(`changes=${saveSummary.changesCount}`);
    }
    if (saveSummary.changeBytes != null) {
      extraParts.push(`bytes=${saveSummary.changeBytes}`);
    }
    if (saveSummary.start || saveSummary.end) {
      extraParts.push(
        `${saveSummary.start ? "start" : ""}${saveSummary.end ? "+end" : ""}`
      );
    }
    if (saveSummary.isExcel) {
      extraParts.push("excel");
    }
    if (saveSummary.deleteIndex != null) {
      extraParts.push(`deleteIndex=${saveSummary.deleteIndex}`);
    }
    if (saveSummary.textPreview.length > 0) {
      extraParts.push(`text="${saveSummary.textPreview.join(" | ")}"`);
    }
  } else if (type === "getLock" && record) {
    const lockSummary = summarizeLockBlock(record.block);
    extra = lockSummary;
    extraParts.push(`locks=${lockSummary.lockCount}`);
    const firstLock = lockSummary.locks[0];
    if (firstLock?.guid) {
      extraParts.push(`guid=${firstLock.guid}`);
    } else if (firstLock?.raw) {
      extraParts.push(`block=${firstLock.raw}`);
    }
  } else if ((type === "getLock" || type === "releaseLock") && record?.locks && typeof record.locks === "object") {
    const lockKeys = Object.keys(record.locks);
    extra = { lockCount: lockKeys.length, lockKeys: lockKeys.slice(0, 3) };
    extraParts.push(`locks=${lockKeys.length}`);
    if (lockKeys[0]) extraParts.push(`key=${truncateText(lockKeys[0], 32)}`);
  } else if (type === "unSaveLock" && record) {
    const index = getNumberField(record, "index");
    const syncChangesIndex = getNumberField(record, "syncChangesIndex");
    extra = { index, syncChangesIndex };
    if (index != null) extraParts.push(`index=${index}`);
    if (syncChangesIndex != null) {
      extraParts.push(`sync=${syncChangesIndex}`);
    }
  }
  const labelParts = [
    type ?? typeof first,
    dataType,
    getStringField(data, "status")
  ].filter(Boolean);
  const label = [
    labelParts.join(" / "),
    extraParts.length > 0 ? `(${extraParts.join(", ")})` : ""
  ].filter(Boolean).join(" ");
  return {
    label,
    type,
    dataType,
    status: getStringField(data, "status"),
    payloadCount: payloads.length,
    ...extra
  };
}
function summarizeDownloadCommand(cmd, buffer) {
  const outputFormat = getNumberField(cmd, "outputformat");
  const filetype = outputFormat == null ? void 0 : extensionFromOutputFormat(outputFormat);
  const saveType = getSaveTypeLabel(cmd.savetype);
  return {
    label: [
      saveType,
      filetype ? `to ${filetype}` : void 0,
      cmd.isSaveAs ? "saveAs" : void 0
    ].filter(Boolean).join(" "),
    saveType,
    outputFormat,
    filetype,
    title: getStringField(cmd, "title"),
    byteLength: buffer.byteLength
  };
}
var EditorServer = class {
  id = "";
  sockets = /* @__PURE__ */ new Set();
  messageHandlers = /* @__PURE__ */ new WeakMap();
  sessionId = "session-id";
  user = {
    id: "uid",
    name: "Me"
  };
  client = {
    buildVersion: "9.4.0",
    buildNumber: 129
  };
  participants = [];
  syncChangesIndex = 0;
  loadPromise = null;
  loadBlocked = false;
  fileType = "docx";
  title = "";
  fsMap = /* @__PURE__ */ new Map();
  urlsMap = /* @__PURE__ */ new Map();
  downloadId = "";
  /**
   * @description downloadAs multipart 分片缓冲；保存与导出共用 HTTP 管道，需与 pendingExport 配合区分意图。
   */
  downloadParts = [];
  /**
   * @description 当前 downloadAs 请求的 cmd（含 outputformat / isSaveAs / title）。
   */
  downloadCmd = null;
  /**
   * @description 另存为 GET 响应 Content-Disposition 使用的文件名。
   */
  downloadFileNames = /* @__PURE__ */ new Map();
  /**
   * @description 用户保存 downloadAs 进行中时阻塞 export，避免分片交错污染 Editor.bin。
   */
  savingDone = Promise.resolve();
  finishSaving = null;
  /**
   * @description export() 调用 downloadAs("bin") 后等待 resolvePendingExport 完成。
   */
  pendingExport = null;
  options = {};
  officeXmlEventConfig = {
    ...OFFICE_XML_EVENT_CONFIG.default
  };
  constructor(options = {}) {
    this.options = options;
    this.handleConnect = this.handleConnect.bind(this);
    this.handleMessage = this.handleMessage.bind(this);
  }
  logRaw(level, category, message, consoleArgs, ...details) {
    const logger = this.options.logger;
    if (logger) {
      logger.raw(level, category, message, consoleArgs, ...details);
      return;
    }
    console[level](...consoleArgs);
  }
  setOfficeXmlEventConfig(config) {
    this.officeXmlEventConfig = {
      ...OFFICE_XML_EVENT_CONFIG.default,
      ...config
    };
  }
  createLoadPromise(buffer, fileType) {
    return this.loadDocument(buffer, fileType).catch((err) => {
      const error = err instanceof Error ? err : new Error(String(err));
      this.notifyLoadError(error);
      if (isOfficeXmlSizeLimitExceededError(error)) {
        this.loadBlocked = true;
        return;
      }
      throw error;
    });
  }
  notifyLoadError(error) {
    try {
      this.options.onLoadError?.(error);
    } catch (callbackError) {
      console.error("[EditorServer] onLoadError callback failed", callbackError);
    }
  }
  /**
   * @description CDN iframe 与主站跨域：Editor.bin 必须用相对 /cache/files/，由 bridge 代拉。
   */
  getCacheFileUrl(name) {
    return `/cache/files/data/${this.id}/${name}`;
  }
  setEditorBinUrl(data) {
    this.fsMap.set("Editor.bin", data);
    if (typeof window !== "undefined" && isOnlyOfficeCdnMode()) {
      this.urlsMap.set("Editor.bin", this.getCacheFileUrl("Editor.bin"));
      return;
    }
    this.urlsMap.set("Editor.bin", getUrl(data));
  }
  reset() {
    if (this.pendingExport) {
      window.clearTimeout(this.pendingExport.timer);
      this.pendingExport.reject(new Error("Editor server reset"));
      this.pendingExport = null;
    }
    for (const socket of this.sockets) {
      const handler = this.messageHandlers.get(socket);
      if (handler) {
        socket.server.off("message", handler);
      }
    }
    this.sockets.clear();
    this.messageHandlers = /* @__PURE__ */ new WeakMap();
    if (this.urlsMap.size > 0) {
      this.urlsMap.forEach((url) => URL.revokeObjectURL(url));
    }
    this.id = "";
    this.fileType = "docx";
    this.title = "";
    this.fsMap.clear();
    this.urlsMap.clear();
    this.loadPromise = null;
    this.loadBlocked = false;
    this.downloadId = "";
    this.downloadParts = [];
    this.downloadCmd = null;
    this.downloadFileNames.clear();
    this.endSaving();
    this.syncChangesIndex = 0;
    this.participants = [];
  }
  async open(file, { fileType, fileName } = {}) {
    this.fileType = fileType || getFileExt(file.name) || "docx";
    const title = ensureTitleWithExtension(fileName || file.name, this.fileType);
    const documentType = getDocumentType(this.fileType);
    this.id = randomId();
    this.title = title;
    const buffer = await file.arrayBuffer();
    const sizeLimitError = this.getOfficeXmlSizeLimitError(
      buffer,
      this.fileType
    );
    if (sizeLimitError) {
      this.loadBlocked = true;
      this.loadPromise = Promise.resolve();
      this.notifyLoadError(sizeLimitError);
      return {
        id: this.id,
        documentType
      };
    }
    this.loadPromise = this.createLoadPromise(buffer, this.fileType);
    return {
      id: this.id,
      documentType
    };
  }
  openNew(fileType, fileName) {
    this.fileType = fileType || "docx";
    this.id = randomId();
    this.loadPromise = null;
    this.title = ensureTitleWithExtension(
      fileName || "New Document",
      this.fileType
    );
    const documentType = getDocumentType(this.fileType);
    let binData = null;
    switch (documentType) {
      case "word":
        binData = Uint8Array.from(emptyDocx, (v) => v.charCodeAt(0));
        break;
      case "cell":
        binData = Uint8Array.from(emptyXlsx, (v) => v.charCodeAt(0));
        break;
      case "slide":
        binData = Uint8Array.from(emptyPptx, (v) => v.charCodeAt(0));
        break;
      case "pdf":
        binData = Uint8Array.from(emptyPdf, (v) => v.charCodeAt(0));
        break;
    }
    if (!binData) {
      throw new Error("Failed to create new document");
    }
    this.fsMap.set("Editor.bin", binData);
    this.setEditorBinUrl(binData);
    this.loadBlocked = false;
    return {
      id: this.id,
      documentType
    };
  }
  async openUrl(url, {
    fileType,
    fileName,
    loader = (url2) => fetch(url2).then((res) => res.arrayBuffer())
  } = {}) {
    const rawTitle = fileName || decodeURIComponent(url.split("/").pop() || "Document");
    this.fileType = fileType || getFileExt(rawTitle) || "docx";
    const documentType = getDocumentType(this.fileType);
    this.id = randomId();
    this.loadBlocked = false;
    this.title = ensureTitleWithExtension(rawTitle, this.fileType);
    this.loadPromise = this.createLoadPromise(() => loader(url), this.fileType);
    return {
      id: this.id,
      documentType
    };
  }
  getDocument() {
    if (!this.id) {
      this.openNew();
    }
    return {
      fileType: this.fileType,
      key: this.id,
      title: this.title,
      url: this.urlsMap.get("Editor.bin") || this.getCacheFileUrl("Editor.bin")
    };
  }
  /**
   * @description 更新当前内存文档标题。文件内容与 document key 保持不变。
   */
  rename(fileName) {
    const requestedName = fileName.trim();
    if (!requestedName) {
      throw new Error("OnlyOffice document name cannot be empty");
    }
    this.title = ensureTitleWithExtension(requestedName, this.fileType);
    this.options.onDocumentRename?.(this.title);
    return this.title;
  }
  /**
   * @description 另存为产物 blob URL，cache GET 失败时供 getFile 回退下载。
   */
  getStoredOutputUrl(outputName) {
    return this.urlsMap.get(outputName) ?? null;
  }
  getStoredOutputFileName(outputName) {
    return this.downloadFileNames.get(outputName) ?? null;
  }
  getUser() {
    return { ...this.user };
  }
  setUser(user) {
    this.user = {
      ...this.user,
      ...user
    };
  }
  getDocumentSnapshot() {
    const binData = this.fsMap.get("Editor.bin");
    const media = this.getStoredMedia();
    const themes = this.getStoredThemes();
    return {
      fileName: this.title,
      fileType: this.fileType,
      binData,
      media,
      themes
    };
  }
  getStoredMedia() {
    return Object.fromEntries(
      Array.from(this.fsMap.entries()).filter(
        ([key]) => key.startsWith("media/")
      )
    );
  }
  getStoredThemes() {
    return Object.fromEntries(
      Array.from(this.fsMap.entries()).filter(
        ([key]) => key.startsWith("themes/")
      )
    );
  }
  getExportSafeMedia() {
    const media = {};
    for (const [key, value] of Object.entries(this.getStoredMedia())) {
      if (/\.(emf|svg|webp)$/i.test(key)) {
        media[key.replace(/\.(emf|svg|webp)$/i, ".png")] = TRANSPARENT_PNG;
        continue;
      }
      media[key] = value;
    }
    return media;
  }
  getExportSafeEditorBin(binData) {
    const hasUnsupportedVectorMedia = Object.keys(this.getStoredMedia()).some(
      (key) => /\.(emf|svg)$/i.test(key)
    );
    if (!hasUnsupportedVectorMedia) {
      return binData;
    }
    return rewriteEditorBinUnsupportedMediaRefs(binData);
  }
  /**
   * @description 用户保存：更新 Editor.bin 并通知接入层，不触发浏览器下载。
   */
  commitUserSave(data) {
    if (!isValidEditorBin(data)) {
      console.warn(
        "[EditorServer] Ignoring invalid Editor.bin from save, length:",
        data.length
      );
      return;
    }
    this.downloadParts = [];
    this.downloadId = "";
    this.updateEditorBin(data);
    this.options.onUserSave?.(this.getDocumentSnapshot());
  }
  beginSaving() {
    this.savingDone = new Promise((resolve) => {
      this.finishSaving = resolve;
    });
  }
  endSaving() {
    this.finishSaving?.();
    this.finishSaving = null;
    this.savingDone = Promise.resolve();
  }
  /**
   * @description 捕获当前文档快照；通过 downloadAs("bin") 让 iframe 把最新 Editor.bin 回传到 mock 服务。
   */
  async captureCurrentDocument(trigger, timeout = 3e4) {
    await this.savingDone;
    this.downloadParts = [];
    this.downloadId = "";
    this.downloadCmd = null;
    if (this.pendingExport) {
      window.clearTimeout(this.pendingExport.timer);
      this.pendingExport.reject(
        new DOMException("OnlyOffice export was superseded", "AbortError")
      );
      this.pendingExport = null;
    }
    return new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        this.pendingExport = null;
        reject(new Error("Timed out waiting for OnlyOffice export data"));
      }, timeout);
      this.pendingExport = { resolve, reject, timer };
      try {
        trigger();
      } catch (err) {
        window.clearTimeout(timer);
        this.pendingExport = null;
        reject(err instanceof Error ? err : new Error("Failed to start export"));
      }
    });
  }
  /**
   * @description 打开文档：x2t 将 doc.{fileType} 转为 Editor.bin 写入 fsMap，供 iframe 加载。
   */
  async loadDocument(buffer, fileType) {
    if (typeof buffer == "function") {
      buffer = await buffer();
    }
    const bytes = new Uint8Array(buffer);
    const fallback = isOfficeZipFileType(fileType) && !isZipBytes(bytes) ? createPlainTextOfficeFallback(buffer, fileType) : null;
    if (fallback) {
      buffer = fallback.buffer;
    }
    const sourceFileType = fallback?.fileType ?? fileType;
    let output = null;
    let media = {};
    let themes = {};
    if (sourceFileType == "pdf") {
      output = new Uint8Array(buffer);
    } else if (sourceFileType === "csv") {
      ({ output, media } = await this.loadCsvDocument(buffer));
    } else {
      ({ output, media, themes } = await this.convertBufferToEditorBin(
        buffer,
        sourceFileType
      ));
    }
    if (!output) {
      throw new Error(`Failed to convert ${sourceFileType} file`);
    }
    if (this.urlsMap.size > 0) {
      this.urlsMap.forEach((url) => URL.revokeObjectURL(url));
    }
    this.fsMap.set("Editor.bin", output);
    this.setEditorBinUrl(output);
    for (const name in media) {
      this.addMedia(name, media[name]);
    }
    for (const [name, data] of Object.entries(themes)) {
      this.fsMap.set(name, data);
    }
  }
  async convertBufferToEditorBin(buffer, fileType) {
    this.assertOfficeXmlSizeWithinLimit(buffer, fileType);
    buffer = await this.rewriteUnsupportedPptxImages(buffer, fileType);
    const { formatFrom, formatTo } = getX2tConvertFormats(fileType);
    const result = await converter.convert(
      {
        data: buffer,
        fileFrom: "doc." + fileType,
        fileTo: "Editor.bin",
        formatFrom,
        formatTo
      },
      this.options.logger
    );
    return {
      output: result.output,
      media: result.media,
      themes: result.themes ?? {}
    };
  }
  assertOfficeXmlSizeWithinLimit(buffer, fileType) {
    const error = this.getOfficeXmlSizeLimitError(buffer, fileType);
    if (error) {
      throw error;
    }
  }
  getOfficeXmlSizeLimitError(buffer, fileType) {
    if (!this.officeXmlEventConfig.isEnable) {
      return null;
    }
    if (!isOfficeZipFileType(fileType)) {
      return null;
    }
    const entries2 = readZipEntries(buffer);
    if (!entries2) {
      return null;
    }
    const { totalSize, entryCount } = getZipXmlUncompressedSize(entries2);
    this.logRaw("error", "operation", "office xml total size", [
      "onlyoffice-totalSize:",
      `${(totalSize / 1024 / 1024).toFixed(2)} MB`
    ]);
    if (totalSize <= this.officeXmlEventConfig.limitBytes) {
      return null;
    }
    return new OfficeXmlSizeLimitExceededError({
      fileName: this.title || `Document.${getFileExt(fileType) || fileType}`,
      fileType: getFileExt(fileType) || fileType,
      errorDescription: "文件过大，不支持解析",
      xmlBytes: totalSize,
      limitBytes: this.officeXmlEventConfig.limitBytes,
      entryCount
    });
  }
  async rewriteUnsupportedPptxImages(buffer, fileType) {
    if (getFileExt(fileType) !== "pptx") {
      return buffer;
    }
    const entries2 = readZipEntries(buffer);
    if (!entries2) {
      return buffer;
    }
    const replacements = /* @__PURE__ */ new Map();
    const entryReplacements = /* @__PURE__ */ new Map();
    const mediaFiles = entries2.filter(
      (entry) => !entry.name.endsWith("/") && /^ppt\/media\/[^/]+\.(emf|svg|webp)$/i.test(entry.name)
    );
    for (const entry of mediaFiles) {
      const pngName = entry.name.replace(/\.(emf|svg|webp)$/i, ".png");
      const source = await readZipEntryData(entry);
      if (!source) {
        continue;
      }
      const pngData = /\.svg$/i.test(entry.name) ? await rasterizeSvgToPng(source) : /\.webp$/i.test(entry.name) ? await rasterizeWebpToPng(source) : TRANSPARENT_PNG;
      entryReplacements.set(pngName, {
        data: pngData,
        modTime: entry.modTime,
        modDate: entry.modDate
      });
      replacements.set(entry.name, pngName);
    }
    if (replacements.size === 0) {
      return buffer;
    }
    const relFiles = entries2.filter(
      (entry) => !entry.name.endsWith("/") && entry.name.endsWith(".rels")
    );
    for (const relFile of relFiles) {
      const source = await readZipEntryData(relFile);
      if (!source) {
        continue;
      }
      let text = new TextDecoder().decode(source);
      let changed = false;
      for (const [from, to] of replacements) {
        const fromName = from.split("/").pop();
        const toName = to.split("/").pop();
        if (!fromName || !toName) {
          continue;
        }
        const nextText = text.replace(
          new RegExp(escapeRegExp(fromName), "g"),
          toName
        );
        if (nextText !== text) {
          text = nextText;
          changed = true;
        }
      }
      if (changed) {
        entryReplacements.set(relFile.name, {
          data: new TextEncoder().encode(text),
          modTime: relFile.modTime,
          modDate: relFile.modDate
        });
      }
    }
    const contentTypes = entries2.find(
      (entry) => entry.name === "[Content_Types].xml"
    );
    if (contentTypes) {
      const source = await readZipEntryData(contentTypes);
      if (source) {
        entryReplacements.set(contentTypes.name, {
          data: new TextEncoder().encode(
            ensurePngContentType(new TextDecoder().decode(source))
          ),
          modTime: contentTypes.modTime,
          modDate: contentTypes.modDate
        });
      }
    }
    return writeZipEntries(entries2, entryReplacements);
  }
  async loadCsvDocument(buffer) {
    if (isMultilineCsv(buffer)) {
      return this.convertCsvViaXlsx(buffer);
    }
    const convertBuffer = sanitizeCsvBufferForX2t(buffer);
    const { formatFrom, formatTo } = getX2tConvertFormats("csv");
    try {
      const result = await converter.convert(
        {
          data: convertBuffer,
          fileFrom: "doc.csv",
          fileTo: "Editor.bin",
          formatFrom,
          formatTo,
          ...getX2tCsvConvertOptions(convertBuffer)
        },
        this.options.logger
      );
      if (result.output?.byteLength) {
        return { output: result.output, media: result.media };
      }
    } catch (err) {
      console.warn("[EditorServer] CSV x2t failed, retry via xlsx:", err);
    }
    return this.convertCsvViaXlsx(buffer);
  }
  async convertCsvViaXlsx(buffer) {
    const xlsxBuffer = await convertCsvBufferToXlsxBuffer(buffer);
    return this.convertBufferToEditorBin(xlsxBuffer, "xlsx");
  }
  addMedia(name, data, mimeType) {
    const pathname = "media/" + name;
    this.fsMap.set(pathname, data);
    const url = getDataUrl(data, mimeType || detectImageMime(data));
    this.urlsMap.set(pathname, url);
    return url;
  }
  updateEditorBin(data) {
    this.setEditorBinUrl(data);
  }
  /**
   * @description 与 Document Server 一致：/cache/files/data/{id}/output.{ext}（站点根绝对路径，便于 iframe 代理命中）
   */
  buildDownloadFileUrl(outputName, downloadFileName) {
    const params = new URLSearchParams({ filename: downloadFileName });
    return `/cache/files/data/${this.id}/${outputName}?${params}`;
  }
  /**
   * @description 另存为：写入 fsMap，返回带文件名的 HTTP 路径（非 blob URL）。
   */
  storeDownloadOutput(outputName, data, downloadFileName) {
    this.fsMap.set(outputName, data);
    this.downloadFileNames.set(outputName, downloadFileName);
    const mimeType = getOnlyOfficeMimeType(
      getFileExt(outputName.replace(/^output\./, ""))
    );
    this.urlsMap.set(outputName, getUrl(data, mimeType));
    return this.buildDownloadFileUrl(outputName, downloadFileName);
  }
  /**
   * @description 判断 downloadAs 请求是否属于「另存为」输出；该路径通常需要服务端 x2t 转换。
   */
  isDownloadAsOutput(cmd, input) {
    if (cmd?.isSaveAs === true) {
      return true;
    }
    if (!isValidEditorBin(input)) {
      return true;
    }
    const outputFormat = typeof cmd?.outputformat === "number" ? cmd.outputformat : void 0;
    if (outputFormat == null || isCanvasBinOutputFormat(outputFormat)) {
      return false;
    }
    const targetExt = extensionFromOutputFormat(outputFormat);
    return Boolean(targetExt && targetExt !== "bin");
  }
  async convertEditorBinToOutput(binData, filetype) {
    binData = this.getExportSafeEditorBin(binData);
    const media = this.getExportSafeMedia();
    const themes = this.getStoredThemes();
    const { formatFrom, formatTo } = getX2tExportFormats(filetype, this.fileType);
    const result = await converter.convert(
      {
        data: binData.slice().buffer,
        fileFrom: "Editor.bin",
        fileTo: `doc.${filetype}`,
        formatFrom,
        formatTo,
        media,
        themes
      },
      this.options.logger
    );
    if (!result.output?.byteLength) {
      throw new Error(`Failed to convert Editor.bin to ${filetype}`);
    }
    return result.output;
  }
  /**
   * @description Web 表格 PDF 另存为 POST 的是渲染器 Memory 流，需回退 fsMap 中的 Editor.bin。
   */
  resolveEditorBinSource(input) {
    if (isValidEditorBin(input)) {
      return input;
    }
    const cached = this.fsMap.get("Editor.bin");
    if (cached && isValidEditorBin(cached)) {
      return cached;
    }
    return null;
  }
  /** 将 Editor.bin 和 Web SDK 渲染器 Memory 流组合导出为 PDF。 */
  async convertEditorBinToPdf(binData, pdfRendererStream) {
    binData = this.getExportSafeEditorBin(binData);
    const media = this.getExportSafeMedia();
    const themes = this.getStoredThemes();
    const { formatFrom, formatTo } = getX2tExportFormats("pdf", this.fileType);
    if (pdfRendererStream?.byteLength) {
      const result2 = await converter.convert(
        {
          data: binData.slice().buffer,
          fileFrom: "output.bin",
          fileTo: "output.pdf",
          media,
          themes,
          pdfBin: pdfRendererStream
        },
        this.options.logger
      );
      if (!result2.output?.byteLength) {
        throw new Error("Failed to convert Editor.bin + pdf.bin to PDF");
      }
      assertValidPdfOutput(new Uint8Array(result2.output), "pdf.bin");
      return result2.output;
    }
    const result = await converter.convert(
      {
        data: binData.slice().buffer,
        fileFrom: "Editor.bin",
        fileTo: "doc.pdf",
        formatFrom,
        formatTo,
        media,
        themes
      },
      this.options.logger
    );
    if (!result.output?.byteLength) {
      throw new Error("Failed to convert Editor.bin to PDF");
    }
    assertValidPdfOutput(new Uint8Array(result.output));
    return result.output;
  }
  async resolveDownloadOutputData(input, filetype) {
    const ext = getFileExt(filetype);
    if (ext === "pdf" && isPdfBytes(input)) {
      return input;
    }
    if (ext === "pdf") {
      const binSource = this.resolveEditorBinSource(input);
      if (!binSource) {
        throw new Error("PDF export failed: document snapshot unavailable");
      }
      const pdfRendererStream = isValidEditorBin(input) ? void 0 : input;
      try {
        return new Uint8Array(
          await withTimeout(
            this.convertEditorBinToPdf(binSource, pdfRendererStream),
            PDF_CONVERT_TIMEOUT_MS,
            "PDF conversion"
          )
        );
      } catch (err) {
        throw err instanceof Error ? err : new Error("PDF export failed", { cause: err });
      }
    }
    if (!isValidEditorBin(input)) {
      return input;
    }
    return new Uint8Array(
      await this.convertEditorBinToOutput(input, filetype)
    );
  }
  resolveDownloadFileType(cmd) {
    const fromOutput = extensionFromOutputFormat(
      typeof cmd?.outputformat === "number" ? cmd.outputformat : void 0
    );
    if (fromOutput) {
      return normalizeX2tExportFileType(fromOutput);
    }
    const title = typeof cmd?.title === "string" ? cmd.title : "";
    return normalizeX2tExportFileType(getFileExt(title) || this.fileType);
  }
  resolveDownloadFileName(cmd, filetype) {
    const title = typeof cmd?.title === "string" ? cmd.title.trim() : "";
    if (title) {
      return ensureTitleWithExtension(title, filetype);
    }
    return ensureTitleWithExtension(this.title, filetype);
  }
  resolvePendingExport(data) {
    const pendingExport = this.pendingExport;
    if (!pendingExport) return false;
    window.clearTimeout(pendingExport.timer);
    this.pendingExport = null;
    if (!isValidEditorBin(data)) {
      pendingExport.reject(
        new Error("OnlyOffice export returned invalid document data")
      );
      return true;
    }
    this.updateEditorBin(data);
    pendingExport.resolve(this.getDocumentSnapshot());
    return true;
  }
  setClient(info) {
    this.client = {
      ...this.client,
      ...info
    };
  }
  registerSocketTransport(socket) {
    if (this.sockets.has(socket)) {
      return;
    }
    this.sockets.add(socket);
    const handler = (msg, ...args) => {
      void this.handleMessage(socket, msg, ...args);
    };
    this.messageHandlers.set(socket, handler);
    socket.server.on("message", handler);
  }
  sendCoAuthoringHandshake(socket) {
    const { sessionId, client } = this;
    const readOnly = this.options.getState?.()?.readOnly ?? false;
    this.participants = [
      {
        connectionId: this.sessionId,
        encrypted: false,
        id: this.user.id,
        idOriginal: this.user.id,
        indexUser: 1,
        isCloseCoAuthoring: false,
        isLiveViewer: readOnly,
        username: this.user.name,
        view: readOnly
      }
    ];
    this.sendTo(socket, {
      maxPayload: 1e8,
      pingInterval: 25e3,
      pingTimeout: 2e4,
      sid: sessionId,
      upgrades: []
    });
    this.sendTo(socket, {
      type: "license",
      license: {
        type: 3,
        buildNumber: client.buildNumber,
        buildVersion: client.buildVersion,
        light: false,
        mode: 0,
        rights: 1,
        protectionSupport: true,
        isAnonymousSupport: true,
        liveViewerSupport: true,
        branding: false,
        customization: true,
        // Developer Edition Connector 通过 advancedApi 门控；关闭时编辑器会
        // 静默忽略 createConnector() 的 postMessage 请求。
        advancedApi: true
      }
    });
  }
  handleConnect({ socket }) {
    this.logRaw("log", "socket", "connect", ["connect: ", socket]);
    this.registerSocketTransport(socket);
    this.sendCoAuthoringHandshake(socket);
  }
  handleDisconnect({ socket }) {
    this.logRaw("log", "socket", "disconnect", ["disconnect: ", socket]);
    const handler = this.messageHandlers.get(socket);
    if (handler) {
      socket.server.off("message", handler);
      this.messageHandlers.delete(socket);
    }
    this.sockets.delete(socket);
  }
  async handleImgUrls(socket, message) {
    const send = (...payload) => this.sendTo(socket, ...payload);
    const images = Array.isArray(message.data) ? message.data : [];
    const urls = [];
    for (let i = 0; i < images.length; i++) {
      const item = images[i];
      if (typeof item !== "string") {
        continue;
      }
      let parsed = parseClipboardImage(item);
      if (!parsed) {
        try {
          parsed = await fetchClipboardImage(item);
        } catch (err) {
          console.warn("[EditorServer] failed to fetch image url:", item, err);
        }
      }
      if (!parsed) {
        continue;
      }
      const filename = `${Date.now()}_${i}.${parsed.ext}`;
      const pathname = `media/${filename}`;
      const url = this.addMedia(filename, parsed.data, parsed.mime);
      urls.push({ path: pathname, url });
    }
    send({
      type: "documentOpen",
      data: {
        type: "imgurls",
        status: "ok",
        /**
         * @description SDK 用 $dc(error) 判断是否弹错；缺省 undefined 会变成未知错误。
         */
        data: { urls, error: 0 }
      }
    });
  }
  sendTo(socket, ...msg) {
    const summary = summarizeWsPayload(msg);
    this.logRaw(
      "log",
      "socket",
      `ws outgoing: ${summary.label}`,
      ["[ws] >> ", ...msg],
      summary
    );
    socket.server.emit("message", ...msg);
  }
  broadcast(...msg) {
    for (const socket of this.sockets) {
      this.sendTo(socket, ...msg);
    }
  }
  async handleMessage(socket, msg, ...args) {
    const incomingSummary = summarizeWsPayload([msg, ...args]);
    this.logRaw(
      "log",
      "socket",
      `ws incoming: ${incomingSummary.label}`,
      ["[ws] << ", msg, args],
      incomingSummary
    );
    const send = (...payload) => this.sendTo(socket, ...payload);
    const { sessionId, participants, user, client } = this;
    const type = typeof msg === "object" && msg && "type" in msg ? msg.type : null;
    switch (type) {
      case "auth": {
        const changes = [];
        const readOnly = this.options.getState?.()?.readOnly ?? false;
        send({
          type: "authChanges",
          changes
        });
        send({
          type: "auth",
          result: 1,
          sessionId,
          participants,
          locks: [],
          indexUser: 1,
          buildVersion: client.buildVersion || "9.4.0",
          buildNumber: client.buildNumber || 9,
          licenseType: 3,
          editorType: 2,
          mode: readOnly ? "view" : "edit",
          permissions: {
            comment: true,
            chat: true,
            download: true,
            edit: !readOnly,
            fillForms: false,
            modifyFilter: !readOnly,
            protect: !readOnly,
            print: true,
            review: true,
            copy: true
          }
        });
        try {
          if (this.loadPromise) {
            await this.loadPromise;
          }
          if (this.loadBlocked) {
            break;
          }
          send({
            type: "documentOpen",
            data: {
              type: "open",
              status: "ok",
              data: {
                ...Object.fromEntries(this.urlsMap)
              }
            }
          });
        } catch (err) {
          console.error(err);
          const message = err instanceof Error ? err.message : String(err);
          if (isOfficeXmlSizeLimitExceededError(err)) {
            break;
          }
          send({
            type: "documentOpen",
            data: {
              type: "open",
              status: "err",
              data: {
                error: message,
                message
              }
            }
          });
        }
        break;
      }
      case "isSaveLock":
        send({
          type: "saveLock",
          saveLock: false
        });
        break;
      case "saveChanges":
        send({
          type: "unSaveLock",
          index: typeof msg.startSaveChanges === "number" ? msg.startSaveChanges : typeof msg.endSaveChanges === "number" ? msg.endSaveChanges : -1,
          syncChangesIndex: ++this.syncChangesIndex,
          time: +/* @__PURE__ */ new Date()
        });
        break;
      case "getLock":
        if (msg.block == null) {
          break;
        }
        {
          const locks = buildCoAuthoringLocks(
            msg.block,
            this.fileType,
            user?.id
          );
          send({ type: "getLock", locks });
          send({ type: "releaseLock", locks });
        }
        break;
      case "openDocument": {
        const message = msg.message;
        if (message?.c === "imgurls") {
          void this.handleImgUrls(socket, message);
        }
        break;
      }
      case "rpc": {
        const rpc = msg.data;
        if (rpc?.type !== "wopi_RenameFile") {
          break;
        }
        const responseKey = msg.responseKey;
        if (typeof responseKey !== "number") {
          break;
        }
        try {
          const title = this.rename(String(rpc.name ?? ""));
          const extension = getFileExt(title);
          const name = extension ? title.slice(0, -(extension.length + 1)) : title;
          send({ type: "rpc", responseKey, data: { Name: name } });
        } catch (error) {
          this.logRaw("warn", "socket", "wopi rename failed", [
            "[EditorServer] wopi rename failed:",
            error
          ]);
          send({ type: "rpc", responseKey, data: null });
        }
        break;
      }
    }
  }
  /**
   * @description Mock 协作 HTTP 入口，OnlyOffice iframe 内 XHR/fetch 会被代理到这里。
   */
  async handleRequest(req) {
    const u = new URL(req.url);
    const { id: key } = this;
    const cacheMatch = u.pathname.match(/\/cache\/files\/data\/([^/]+)\/(.+)$/);
    if (cacheMatch && req.method === "GET") {
      const [, , rawName] = cacheMatch;
      const outputName = decodeURIComponent(rawName.split("?")[0]);
      if (!this.fsMap.has(outputName) && this.loadPromise) {
        await this.loadPromise;
      }
      const data = this.fsMap.get(outputName);
      if (!data) {
        return new Response("Not Found", { status: 404 });
      }
      const filetype = getFileExt(outputName.replace(/^output\./, "")) || this.fileType;
      const downloadName = getCacheDownloadName(
        outputName,
        this.title,
        filetype,
        this.downloadFileNames
      );
      const mimeType = getCacheResponseMimeType(outputName, this.fileType);
      return new Response(data, {
        status: 200,
        headers: {
          "Content-Type": mimeType,
          "Content-Disposition": buildContentDisposition(downloadName),
          "Content-Length": String(data.byteLength)
        }
      });
    }
    if (u.pathname.endsWith("/downloadas/" + key)) {
      const cmd = JSON.parse(u.searchParams.get("cmd") || "{}");
      const buffer = await req.arrayBuffer();
      const downloadSummary = summarizeDownloadCommand(cmd, buffer);
      this.logRaw(
        "log",
        "download",
        `downloadAs chunk: ${downloadSummary.label || "unknown"}`,
        ["downloadAs -> ", cmd, buffer],
        downloadSummary
      );
      if (cmd.savetype === 0 /* PartStart */ || cmd.savetype === 3 /* CompleteAll */) {
        this.downloadCmd = cmd;
      }
      const download = async () => {
        const input = mergeBuffers(this.downloadParts);
        const cmdSnapshot = this.downloadCmd;
        this.downloadCmd = null;
        const resolvedExport = this.resolvePendingExport(input);
        if (resolvedExport) {
          this.endSaving();
          return {
            status: "ok",
            isExport: true,
            isSaveAs: false,
            downloadUrl: "",
            filetype: "bin"
          };
        }
        if (this.isDownloadAsOutput(cmdSnapshot, input)) {
          const filetype = this.resolveDownloadFileType(cmdSnapshot);
          const outputName = `output.${filetype}`;
          const downloadFileName = this.resolveDownloadFileName(
            cmdSnapshot,
            filetype
          );
          const outputData = await this.resolveDownloadOutputData(
            input,
            filetype
          );
          const downloadUrl = this.storeDownloadOutput(
            outputName,
            outputData,
            downloadFileName
          );
          this.endSaving();
          return {
            status: "ok",
            isExport: false,
            isSaveAs: true,
            downloadUrl,
            filetype
          };
        }
        this.commitUserSave(input);
        this.endSaving();
        return {
          status: "ok",
          isExport: false,
          isSaveAs: false,
          downloadUrl: this.urlsMap.get("Editor.bin") || "",
          filetype: this.fileType
        };
      };
      let result = {
        status: "ok",
        isExport: false,
        isSaveAs: false,
        downloadUrl: "",
        filetype: this.fileType
      };
      let isFinalChunk = false;
      const finalizeDownload = async () => {
        try {
          result = await download();
        } catch (err) {
          this.logRaw("error", "download", "downloadAs failed", [
            "[EditorServer] downloadAs failed:",
            err
          ]);
          this.endSaving();
          result = {
            status: "err",
            isExport: false,
            isSaveAs: false,
            downloadUrl: "",
            filetype: this.fileType,
            error: err instanceof Error ? err.message : String(err)
          };
        }
        this.downloadParts = [];
        isFinalChunk = true;
      };
      switch (cmd.savetype) {
        case 0 /* PartStart */:
          if (!this.pendingExport) {
            this.beginSaving();
          }
          this.downloadId = "_" + Math.round(Math.random() * 1e3);
          this.downloadParts = [new Uint8Array(buffer)];
          break;
        case 1 /* Part */:
          this.downloadParts.push(new Uint8Array(buffer));
          break;
        case 2 /* Complete */:
          this.downloadParts.push(new Uint8Array(buffer));
          await finalizeDownload();
          break;
        case 3 /* CompleteAll */:
          if (!this.pendingExport) {
            this.beginSaving();
          }
          this.downloadId = "_" + Math.round(Math.random() * 1e3);
          this.downloadParts = [new Uint8Array(buffer)];
          await finalizeDownload();
          break;
      }
      if (isFinalChunk) {
        const { downloadUrl, filetype, status, error } = result;
        setTimeout(() => {
          this.broadcast({
            type: "documentOpen",
            data: {
              type: "save",
              status,
              data: status === "ok" ? result.isExport ? PROGRAMMATIC_EXPORT_ACK_URL : downloadUrl : error || "download failed",
              filetype
            }
          });
        }, 100);
      }
      return Response.json({
        status: result.status,
        type: "save",
        data: this.downloadId
      });
    }
    if (u.pathname.endsWith("/upload/" + key)) {
      const buffer = await req.arrayBuffer();
      const data = new Uint8Array(buffer);
      const mime = detectImageMime(data);
      const filename = `${Date.now()}.${getImageExt(mime)}`;
      const pathname = "media/" + filename;
      const url = this.addMedia(filename, data, mime);
      return Response.json({ [pathname]: url });
    }
    return null;
  }
};

// vendor/onlyoffice-web-comp/internal/editor/logger.ts
var EditorLogger = class {
  constructor(editorId) {
    this.editorId = editorId;
  }
  entries = [];
  nextId = 1;
  getConsoleMethod(level) {
    return console[level] ?? console.log;
  }
  write(level, category, message, details, consoleArgs) {
    const entry = {
      id: this.nextId++,
      time: (/* @__PURE__ */ new Date()).toISOString(),
      level,
      category,
      message,
      details
    };
    this.entries.push(entry);
    this.getConsoleMethod(level)(
      ...consoleArgs ?? [
        `[OnlyOffice:${this.editorId}] [${category}] ${message}`,
        ...details
      ]
    );
  }
  log(category, message, ...details) {
    this.write("log", category, message, details);
  }
  info(category, message, ...details) {
    this.write("info", category, message, details);
  }
  warn(category, message, ...details) {
    this.write("warn", category, message, details);
  }
  error(category, message, ...details) {
    this.write("error", category, message, details);
  }
  raw(level, category, message, consoleArgs, ...details) {
    this.write(
      level,
      category,
      message,
      details.length > 0 ? details : consoleArgs,
      consoleArgs
    );
  }
  socket(message, ...details) {
    this.log("socket", message, ...details);
  }
  download(message, ...details) {
    this.log("download", message, ...details);
  }
  worker(message, ...details) {
    this.log("worker", message, ...details);
  }
  operation(message, ...details) {
    this.log("operation", message, ...details);
  }
  getEntries() {
    return [...this.entries];
  }
  clear() {
    this.entries = [];
  }
  print() {
    const title = `[OnlyOffice:${this.editorId}] editor logs (${this.entries.length})`;
    if (typeof console.groupCollapsed === "function") {
      console.groupCollapsed(title);
      const rows = this.entries.map(({ id, time, level, category, message }) => ({
        id,
        time,
        level,
        category,
        message
      }));
      if (typeof console.table === "function") {
        console.table(rows);
      } else {
        console.log(rows);
      }
      this.entries.forEach((entry) => {
        this.getConsoleMethod(entry.level)(
          `#${entry.id} ${entry.time} [${entry.category}] ${entry.message}`,
          ...entry.details
        );
      });
      console.groupEnd();
      return;
    }
    console.log(title, this.getEntries());
  }
};

// vendor/onlyoffice-web-comp/feature/comments.ts
function normalizeCommentInput(input) {
  if (typeof input === "string") {
    return { Text: input };
  }
  return input;
}
function toPluginCommentPayload(data) {
  const now = String(Date.now());
  return {
    Text: data.Text ?? "",
    UserName: data.UserName ?? "Guest",
    Time: data.Time != null ? String(data.Time) : now,
    QuoteText: data.QuoteText ?? "",
    Solved: data.Solved ?? false,
    UserData: data.UserData ?? "",
    Replies: Array.isArray(data.Replies) ? data.Replies : [],
    ...data
  };
}
function isResolvedComment(data) {
  if (!data || typeof data !== "object") {
    return false;
  }
  const comment = data;
  const solved = comment.Solved ?? comment.solved ?? comment.Resolved ?? comment.resolved ?? comment.Done ?? comment.done;
  if (typeof solved === "boolean") {
    return solved;
  }
  const ascGetSolved = comment.asc_getSolved;
  if (typeof ascGetSolved === "function") {
    return !!ascGetSolved.call(data);
  }
  const getSolved = comment.get_Solved;
  if (typeof getSolved === "function") {
    return !!getSolved.call(data);
  }
  return false;
}

// vendor/onlyoffice-web-comp/core/eventbus.ts
var OnlyOfficeEventBus = class {
  emitter = new EventEmitter();
  on(key, handler) {
    this.emitter.on(key, handler);
  }
  off(key, handler) {
    this.emitter.off(key, handler);
  }
  emit(key, data) {
    this.emitter.emit(key, data);
  }
  waitFor(key, timeout = 3e4) {
    return new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        this.off(key, handler);
        reject(new Error(`Timed out waiting for ${key}`));
      }, timeout);
      const handler = (data) => {
        window.clearTimeout(timer);
        this.off(key, handler);
        resolve(data);
      };
      this.on(key, handler);
    });
  }
};
var onlyofficeEventbus = new OnlyOfficeEventBus();

// vendor/onlyoffice-web-comp/feature/revisions.ts
function getAscRevisionEnums(frameWin) {
  const Asc = frameWin?.Asc;
  const fu = Asc?.c_oAscRevisionsChangeType;
  const Bs = Asc?.c_oAscRevisionsMove;
  return {
    TextAdd: fu?.TextAdd,
    TextRem: fu?.TextRem,
    ParaAdd: fu?.ParaAdd,
    ParaRem: fu?.ParaRem,
    TextPr: fu?.TextPr,
    ParaPr: fu?.ParaPr,
    TablePr: fu?.TablePr,
    RowsAdd: fu?.RowsAdd,
    RowsRem: fu?.RowsRem,
    TableRowPr: fu?.TableRowPr,
    MoveMark: fu?.MoveMark,
    Unknown: fu?.Unknown,
    MoveTo: Bs?.MoveTo,
    MoveFrom: Bs?.MoveFrom,
    NoMove: Bs?.NoMove
  };
}
function revisionTypeName(type, enums) {
  const entries2 = [
    ["TextAdd", enums.TextAdd],
    ["TextRem", enums.TextRem],
    ["ParaAdd", enums.ParaAdd],
    ["ParaRem", enums.ParaRem],
    ["TextPr", enums.TextPr],
    ["ParaPr", enums.ParaPr],
    ["TablePr", enums.TablePr],
    ["RowsAdd", enums.RowsAdd],
    ["RowsRem", enums.RowsRem],
    ["TableRowPr", enums.TableRowPr],
    ["MoveMark", enums.MoveMark],
    ["Unknown", enums.Unknown]
  ];
  for (const [name, val] of entries2) {
    if (val !== void 0 && val === type) return name;
  }
  return String(type);
}
function getLogicDocument(api) {
  if (typeof api.te === "function") return api.te() ?? null;
  if (typeof api.Ge === "function") return api.Ge() ?? null;
  if (typeof api.Gg === "function") return api.Gg() ?? null;
  return null;
}
function revisionStart(raw) {
  if (typeof raw.get_StartPos === "function") {
    const pos = raw.get_StartPos();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.SU === "function") {
    const pos = raw.SU();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.wa === "number") return raw.wa;
  if (typeof raw.j9 === "function") return raw.j9();
  if (typeof raw.sa === "number") return raw.sa;
  return void 0;
}
function revisionEnd(raw) {
  if (typeof raw.get_EndPos === "function") {
    const pos = raw.get_EndPos();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.h5 === "function") {
    const pos = raw.h5();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.xa === "number") return raw.xa;
  if (typeof raw.yza === "function") return raw.yza();
  if (typeof raw.ra === "number") return raw.ra;
  return void 0;
}
function isRevisionGroupedMove(raw) {
  return Array.isArray(raw.OL) && raw.OL.length > 0;
}
function revisionIdentity(raw) {
  if (typeof raw.ed === "function") return String(raw.ed());
  return null;
}
function isSameRevision(a, b) {
  if (a === b) return true;
  const idA = revisionIdentity(a);
  const idB = revisionIdentity(b);
  return idA != null && idA === idB;
}
function resolveFreshRevisionRaw(doc, raw) {
  const id = revisionIdentity(raw);
  if (!id) return raw;
  const qt = doc.Wq?.wih?.();
  if (!qt) return raw;
  for (const group of Object.values(qt)) {
    if (!Array.isArray(group)) continue;
    for (const entry of group) {
      if (revisionIdentity(entry) === id) return entry;
    }
  }
  return raw;
}
function applyRevisionSelection(doc, raw, frameWin) {
  if (isRevisionGroupedMove(raw) && raw.vva) {
    const enums = getAscRevisionEnums(frameWin);
    doc.lc?.();
    doc.Ihc?.(raw.vva, raw.nW === enums.MoveFrom, false, false);
    return true;
  }
  const element = raw.Element ?? (typeof raw.ee === "function" ? raw.ee() : void 0);
  const start = typeof raw.wa === "number" ? raw.wa : revisionStart(raw);
  if (element == null || start == null) return false;
  const end = typeof raw.xa === "number" ? raw.xa : revisionEnd(raw) ?? start;
  try {
    doc.lc?.();
    const el = element;
    if (typeof el.Dp === "function" && typeof el.Gs === "function") {
      if (doc.Wq?.u0d?.([element])) return false;
      el.Dp(start, false, -1, -1);
      if (el.Selection) el.Selection.Na = true;
      el.Gs(start, end, false);
      el.Ft?.(false);
      return true;
    }
    if (typeof el.qp === "function" && typeof el.mo === "function") {
      doc.ec?.();
      el.qp(start, false, -1, -1);
      if (el.Selection) el.Selection.La = true;
      el.mo(start, end, false);
      el.hq?.(false);
      doc.Pf?.(false);
      doc.yf?.(true);
      return true;
    }
    return false;
  } catch {
    return false;
  }
}
function dismissReviewChangesDialog(frameWin) {
  const doc = frameWin?.document;
  if (!doc) return;
  doc.querySelectorAll(".asc-window.review-changes.modal-dlg").forEach((el) => {
    const close = el.querySelector(".close, .btn-close");
    if (close instanceof HTMLElement) {
      close.click();
      return;
    }
    el.style.display = "none";
  });
}
function prepareRevisionReviewDisplay(api, frameWin) {
  api.asc_SetLocalTrackRevisions?.(true);
  if (typeof api.pluginMethod_SetDisplayModeInReview === "function") {
    api.pluginMethod_SetDisplayModeInReview("markup");
    dismissReviewChangesDialog(frameWin);
    return;
  }
  if (typeof api.asc_SetDisplayModeInReview === "function") {
    const Asc = frameWin?.Asc;
    const markup = Asc?.c_oAscReviewDisplay?.Edit ?? Asc?.c_oAscReviewDisplay?.Markup ?? Asc?.Xja?.AKa;
    if (markup !== void 0) {
      api.asc_SetDisplayModeInReview(markup);
    }
  }
  dismissReviewChangesDialog(frameWin);
}
function revisionType(raw) {
  if (typeof raw.get_Type === "function") return raw.get_Type();
  if (typeof raw.Gc === "function") return raw.Gc();
  return void 0;
}
function revisionValue(raw) {
  if (typeof raw.get_Value === "function") {
    const v = raw.get_Value();
    if (typeof v === "string") return v;
    if (v != null) return String(v);
  }
  if (typeof raw.Ym === "function") {
    const v = raw.Ym();
    if (typeof v === "string") return v;
    if (v != null) return String(v);
  }
  return void 0;
}
function revisionChangeToPlain(raw, enums) {
  const type = revisionType(raw);
  const lockUserId = raw.get_LockUserId?.();
  return {
    Type: type,
    TypeName: type !== void 0 ? revisionTypeName(type, enums) : void 0,
    UserId: (typeof raw.get_UserId === "function" ? raw.get_UserId() : void 0) ?? (typeof raw.Qy === "function" ? raw.Qy() : void 0),
    UserName: (typeof raw.get_UserName === "function" ? raw.get_UserName() : void 0) ?? (typeof raw.wV === "function" ? raw.wV() : void 0),
    DateTime: (typeof raw.get_DateTime === "function" ? raw.get_DateTime() : void 0) ?? (typeof raw.nX === "function" ? raw.nX() : void 0),
    Value: revisionValue(raw),
    X: (typeof raw.get_X === "function" ? raw.get_X() : void 0) ?? (typeof raw.$Zd === "function" ? raw.$Zd() : void 0),
    Y: (typeof raw.get_Y === "function" ? raw.get_Y() : void 0) ?? (typeof raw.Vtb === "function" ? raw.Vtb() : void 0),
    MoveType: (typeof raw.get_MoveType === "function" ? raw.get_MoveType() : void 0) ?? (typeof raw.hZd === "function" ? raw.hZd() : void 0),
    MoveId: raw.get_MoveId?.(),
    Locked: lockUserId != null && lockUserId !== "",
    LockUserId: lockUserId ?? void 0
  };
}
var syncingRevisionStack = false;
function ensureRevisionsIndexed(api) {
  const doc = getLogicDocument(api);
  doc?.Wq?.z9a?.();
  doc?.Um?.MQc?.();
}
function collectRevisionRawsFromWq(doc) {
  const wq = doc.Wq;
  const qt = wq?.wih?.();
  if (!qt || typeof qt !== "object" || Object.keys(qt).length === 0) {
    return [];
  }
  const raws = [];
  const orderedIds = [];
  for (const el of wq?.Tpa ?? []) {
    const id = typeof el.Yb === "function" ? String(el.Yb()) : typeof el.ed === "function" ? String(el.ed()) : "";
    if (id && qt[id] && !orderedIds.includes(id)) {
      orderedIds.push(id);
    }
  }
  const rest = Object.keys(qt).filter((id) => !orderedIds.includes(id)).sort((a, b) => {
    const ga = qt[a]?.[0];
    const gb = qt[b]?.[0];
    if (!ga || !gb) return 0;
    return revisionSortKey(ga) - revisionSortKey(gb);
  });
  orderedIds.push(...rest);
  for (const id of orderedIds.length > 0 ? orderedIds : Object.keys(qt)) {
    const group = qt[id];
    if (!Array.isArray(group)) continue;
    raws.push(
      ...group.length > 1 ? [...group].sort((a, b) => revisionSortKey(a) - revisionSortKey(b)) : group
    );
  }
  return raws;
}
function readRevisionStack(api, allowSyncEnd, forceRefresh = false) {
  if (!forceRefresh) {
    const existing = api.asc_GetRevisionsChangesStack?.() ?? [];
    if (existing.length > 0 || !allowSyncEnd || syncingRevisionStack) {
      return existing;
    }
  }
  if (!allowSyncEnd || syncingRevisionStack) {
    return api.asc_GetRevisionsChangesStack?.() ?? [];
  }
  syncingRevisionStack = true;
  try {
    api.sync_BeginCatchRevisionsChanges?.();
    getLogicDocument(api)?.Um?.Dmf?.();
    api.sync_EndCatchRevisionsChanges?.(false);
    return api.asc_GetRevisionsChangesStack?.() ?? [];
  } finally {
    syncingRevisionStack = false;
  }
}
function revisionSortKey(raw) {
  const y = typeof raw.$Zd === "function" ? raw.$Zd() : 0;
  const x = typeof raw.Vtb === "function" ? raw.Vtb() : 0;
  return y * 1e9 + x * 1e3 + (revisionStart(raw) ?? 0);
}
function mapRevisionItems(raws, enums, idOf) {
  return raws.map((raw, index) => ({
    Id: idOf(raw, index),
    Index: index,
    Data: revisionChangeToPlain(raw, enums),
    Raw: raw
  }));
}
function collectRevisionItems(api, frameWin, options = {}) {
  const { allowSyncEnd = true, forceRefreshStack = false } = options;
  ensureRevisionsIndexed(api);
  const enums = getAscRevisionEnums(frameWin);
  const revId = (raw, i) => `rev-${typeof raw.ed === "function" ? String(raw.ed()) : i}`;
  const doc = getLogicDocument(api);
  const wqRaws = doc ? collectRevisionRawsFromWq(doc) : [];
  if (wqRaws.length > 0) {
    return mapRevisionItems(wqRaws, enums, revId);
  }
  const um = doc?.Um;
  const yq = um?.yif?.();
  if (yq && Object.keys(yq).length > 0) {
    const ids = [];
    for (const el of um?.T8 ?? []) {
      const id = typeof el.ed === "function" ? String(el.ed()) : "";
      if (id && yq[id]) ids.push(id);
    }
    const rest = Object.keys(yq).filter((id) => !ids.includes(id));
    rest.sort((a, b) => {
      const ga = yq[a]?.[0];
      const gb = yq[b]?.[0];
      if (!ga || !gb) return 0;
      return revisionSortKey(ga) - revisionSortKey(gb);
    });
    ids.push(...rest);
    const raws = [];
    for (const id of ids.length > 0 ? ids : Object.keys(yq)) {
      const group = yq[id];
      if (!Array.isArray(group)) continue;
      raws.push(
        ...group.length > 1 ? [...group].sort((a, b) => revisionSortKey(a) - revisionSortKey(b)) : group
      );
    }
    if (raws.length > 0) return mapRevisionItems(raws, enums, revId);
  }
  const report = api.asc_GetTrackRevisionsReportByAuthors?.();
  if (report && typeof report === "object") {
    const raws = [];
    for (const key of Object.keys(report)) {
      const group = report[key];
      if (Array.isArray(group)) raws.push(...group);
    }
    if (raws.length > 0) return mapRevisionItems(raws, enums, revId);
  }
  const stack = readRevisionStack(api, allowSyncEnd, forceRefreshStack);
  return mapRevisionItems(stack, enums, (_, i) => `rev-stack-${i}`);
}
function resolveRevisionShowChanges(stack, api, frameWin) {
  if (Array.isArray(stack) && stack.length > 0) {
    const enums = getAscRevisionEnums(frameWin);
    return stack.map((raw, index) => ({
      Id: `rev-stack-${index}`,
      Index: index,
      Data: revisionChangeToPlain(raw, enums),
      Raw: raw
    }));
  }
  return collectRevisionItems(api, frameWin, { allowSyncEnd: false });
}
function findRevisionItem(items, id) {
  return items.find((item) => item.Id === id);
}
function resolveRevisionTarget(target, api, frameWin, cachedItems) {
  if (typeof target !== "string") {
    return target.Raw ? target : void 0;
  }
  const cached = cachedItems?.find((item) => item.Id === target);
  if (cached) return cached;
  return findRevisionItem(
    collectRevisionItems(api, frameWin, { allowSyncEnd: false }),
    target
  );
}
function scrollToRevisionPosition(api, data) {
  if (data.Y != null) api.ViewScrollToY?.(data.Y);
  if (data.X != null) api.ViewScrollToX?.(data.X);
}
function scrollRevisionIntoView(doc, api, data) {
  if (doc.zj?.W5b && doc.zj?.Ptb && typeof doc.qde === "function") {
    doc.qde(null, false);
    doc.pb?.p8a?.();
    return;
  }
  scrollToRevisionPosition(api, data);
}
function activateRevision(doc, raw) {
  const wq = doc.Wq;
  if (!wq) return;
  wq.t0d?.();
  wq.TLh?.(raw);
  wq.Ydf?.(true);
  doc.Ek?.f8?.(wq);
  wq.Znf?.();
  doc.ze?.(true);
}
function navigateToRevisionByStep(api, doc, targetRaw, targetIndex) {
  if (typeof api.asc_GetNextRevisionsChange !== "function") return false;
  doc.Wq?.$_d?.();
  const steps = Math.max(targetIndex + 1, 1);
  for (let i = 0; i < steps; i++) {
    api.asc_GetNextRevisionsChange();
    const current2 = doc.Wq?.kz;
    if (current2 && isSameRevision(current2, targetRaw)) return true;
  }
  const current = doc.Wq?.kz;
  return current != null && isSameRevision(current, targetRaw);
}
function focusRevisionInDocument(api, doc, raw, data, frameWin, index) {
  ensureRevisionsIndexed(api);
  const resolvedRaw = resolveFreshRevisionRaw(doc, raw);
  doc.Wq?.TLh?.(resolvedRaw);
  if (applyRevisionSelection(doc, resolvedRaw, frameWin)) {
    scrollRevisionIntoView(doc, api, data);
    activateRevision(doc, resolvedRaw);
    return true;
  }
  if (navigateToRevisionByStep(api, doc, resolvedRaw, index)) {
    scrollRevisionIntoView(doc, api, data);
    return true;
  }
  scrollToRevisionPosition(api, data);
  activateRevision(doc, resolvedRaw);
  return true;
}
function isMoveRevision(raw, frameWin) {
  const enums = getAscRevisionEnums(frameWin);
  const moveType = (typeof raw.get_MoveType === "function" ? raw.get_MoveType() : void 0) ?? (typeof raw.hZd === "function" ? raw.hZd() : void 0);
  if (moveType === void 0) return false;
  return moveType === enums.MoveTo || moveType === enums.MoveFrom;
}
function goToRevision(target, api, frameWin, cachedItems) {
  const item = resolveRevisionTarget(target, api, frameWin, cachedItems);
  if (!item?.Raw) return false;
  const doc = getLogicDocument(api);
  if (!doc) return false;
  const { Raw: raw, Data: data, Index: index } = item;
  if (isMoveRevision(raw, frameWin)) {
    api.asc_FollowRevisionMove?.(raw);
    scrollToRevisionPosition(api, data);
  }
  return focusRevisionInDocument(api, doc, raw, data, frameWin, index);
}
function applyRevisionChange(mode, target, api, frameWin, cachedItems) {
  const item = resolveRevisionTarget(target, api, frameWin, cachedItems);
  if (!item?.Raw) {
    return false;
  }
  goToRevision(item, api, frameWin, cachedItems);
  const doc = getLogicDocument(api);
  const raw = doc ? resolveFreshRevisionRaw(doc, item.Raw) : item.Raw;
  if (mode === "accept") {
    api.asc_AcceptChanges?.(raw);
  } else {
    api.asc_RejectChanges?.(raw);
  }
  return true;
}

// vendor/onlyoffice-web-comp/internal/ui/office-xml-size-limit-overlay.ts
var OVERLAY_ATTR = "data-onlyoffice-xml-size-limit-overlay";
function formatBytes(bytes) {
  const units = ["B", "KB", "MB", "GB", "TB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}
function removeOfficeXmlSizeLimitOverlay(container) {
  container.querySelector(`[${OVERLAY_ATTR}="true"]`)?.remove();
}
function showOfficeXmlSizeLimitOverlay(container, payload) {
  removeOfficeXmlSizeLimitOverlay(container);
  const computedPosition = window.getComputedStyle(container).position;
  if (computedPosition === "static") {
    container.style.position = "relative";
  }
  const overlay = document.createElement("div");
  overlay.setAttribute(OVERLAY_ATTR, "true");
  overlay.setAttribute("role", "alert");
  overlay.style.position = "absolute";
  overlay.style.inset = "0";
  overlay.style.zIndex = "1";
  overlay.style.display = "flex";
  overlay.style.alignItems = "center";
  overlay.style.justifyContent = "center";
  overlay.style.padding = "24px";
  overlay.style.background = "#f8fafc";
  overlay.style.color = "#172033";
  overlay.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
  const panel = document.createElement("div");
  panel.style.width = "min(520px, 100%)";
  panel.style.border = "1px solid #d8dee8";
  panel.style.borderRadius = "8px";
  panel.style.background = "#ffffff";
  panel.style.boxShadow = "0 16px 42px rgba(15, 23, 42, 0.12)";
  panel.style.padding = "24px";
  const title = document.createElement("div");
  title.textContent = payload.errorDescription || OFFICE_XML_SIZE_LIMIT_ERROR_MESSAGE;
  title.style.fontSize = "18px";
  title.style.lineHeight = "26px";
  title.style.fontWeight = "650";
  title.style.marginBottom = "10px";
  const description = document.createElement("div");
  description.textContent = "Office 文件内 XML 解压后超过安全阈值，请压缩或拆分文档后重试。";
  description.style.fontSize = "14px";
  description.style.lineHeight = "22px";
  description.style.color = "#475569";
  description.style.marginBottom = "16px";
  const details = document.createElement("div");
  details.textContent = `XML 总大小 ${formatBytes(payload.xmlBytes)}，限制 ${formatBytes(payload.limitBytes)}，文件数 ${payload.entryCount}`;
  details.style.fontSize = "12px";
  details.style.lineHeight = "18px";
  details.style.color = "#64748b";
  details.style.wordBreak = "break-word";
  panel.append(title, description, details);
  overlay.append(panel);
  container.append(overlay);
  return () => {
    overlay.remove();
  };
}

// vendor/onlyoffice-web-comp/core/editor-manager.ts
function getFileType(fileName, fileType) {
  return fileType || fileName.split(".").pop()?.toLowerCase() || "docx";
}
var EditorManager = class {
  editor = null;
  /**
   * Connector 和 logger 一样属于一个 EditorManager（也就是一个编辑器 iframe）。
   * 9.4 的 DocsAPI 每创建一个 Connector 都会注册一组 postMessage 监听器；
   * 因此不能由业务调用方重复创建。
   */
  connector = null;
  server;
  dirty = false;
  readOnly = false;
  editorLang = getOnlyOfficeLang();
  uiTheme = "theme-white";
  /** 与容器一一对应，供事件与 Connector 使用同一稳定路由键。 */
  instanceId;
  containerId;
  logger;
  fileName = "New Document.docx";
  fileType = "docx";
  media = {};
  comments = /* @__PURE__ */ new Map();
  crossOriginCommentRefreshPromise = null;
  revisions = [];
  refreshingRevisions = false;
  crossOriginRevisionRefreshPromise = null;
  trackRevisions = false;
  revisionReviewMode = false;
  wordContentSyncPromise = null;
  wordContentSyncTeardown = null;
  crossOriginBridgeTeardown = null;
  officeXmlSizeLimitOverlayTeardown = null;
  officeXmlSizeLimitPayload = null;
  pendingRename = null;
  constructor(containerId = ONLYOFFICE_ID) {
    this.containerId = containerId;
    this.instanceId = containerId;
    this.logger = new EditorLogger(containerId);
    this.server = new EditorServer({
      getState: () => ({ readOnly: this.readOnly }),
      logger: this.logger,
      onUserSave: (snapshot) => {
        this.dirty = false;
        this.notifyUserSave(snapshot);
      },
      onLoadError: (error) => {
        this.handleServerLoadError(error);
      },
      onDocumentRename: (fileName) => {
        this.syncRenamedDocument(fileName);
      }
    });
  }
  getContainerElement() {
    return document.getElementById(this.containerId);
  }
  getOfficeOverlayHostElement() {
    const container = this.getContainerElement();
    return container?.closest(
      ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR
    ) ?? container;
  }
  clearOfficeXmlSizeLimitOverlay() {
    this.officeXmlSizeLimitOverlayTeardown?.();
    this.officeXmlSizeLimitOverlayTeardown = null;
    this.officeXmlSizeLimitPayload = null;
    const container = this.getOfficeOverlayHostElement();
    if (container) {
      removeOfficeXmlSizeLimitOverlay(container);
    }
  }
  renderOfficeXmlSizeLimitOverlay() {
    if (!this.officeXmlSizeLimitPayload) {
      return;
    }
    const container = this.getOfficeOverlayHostElement();
    if (container) {
      this.officeXmlSizeLimitOverlayTeardown = showOfficeXmlSizeLimitOverlay(
        container,
        this.officeXmlSizeLimitPayload
      );
    }
  }
  handleServerLoadError(error) {
    if (!isOfficeXmlSizeLimitExceededError(error)) {
      return;
    }
    this.officeXmlSizeLimitPayload = error.payload;
    this.renderOfficeXmlSizeLimitOverlay();
    const data = {
      ...error.payload,
      instanceId: this.instanceId,
      containerId: this.containerId
    };
    try {
      onlyofficeEventbus.emit(
        ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
        data
      );
    } catch (eventError) {
      console.error(
        "[EditorManager] officeXmlSizeLimitExceeded handler failed",
        eventError
      );
    }
  }
  createScopedIo() {
    return (url, options = {}) => {
      const socket = runtime_bridge_default(url, { ...options, logger: this.logger });
      socket.on("connect", () => {
        this.server.handleConnect({ socket });
      });
      socket.on("disconnect", () => {
        this.server.handleDisconnect({ socket });
      });
      return socket;
    };
  }
  createCrossOriginServerIo() {
    return () => {
      const socket = runtime_bridge_default(void 0, {
        deferConnect: true,
        logger: this.logger
      });
      socket.connected = true;
      socket.disconnected = false;
      return socket;
    };
  }
  syncEditorBridge() {
    this.crossOriginBridgeTeardown?.();
    this.crossOriginBridgeTeardown = null;
    unregisterScopedIo(this.containerId);
    registerScopedIo(this.containerId, this.createScopedIo());
    if (!isOnlyOfficeCdnMode()) {
      return;
    }
    this.crossOriginBridgeTeardown = watchCrossOriginIframe(
      this.containerId,
      () => this.getEditorFrameElement(),
      this.server,
      this.createCrossOriginServerIo()
    );
  }
  syncCrossOriginReadOnly(readOnly, retries = 10) {
    if (setCrossOriginReadOnly(this.containerId, readOnly)) {
      return true;
    }
    if (retries > 0) {
      window.setTimeout(() => {
        this.syncCrossOriginReadOnly(readOnly, retries - 1);
      }, 50);
    }
    return false;
  }
  getEditorFrameElement() {
    const containerFrame = document.getElementById(this.containerId)?.querySelector('iframe[name="frameEditor"]');
    if (containerFrame) {
      return containerFrame;
    }
    const frames = Array.from(
      document.querySelectorAll(
        'iframe[name="frameEditor"]'
      )
    );
    const matchedFrame = frames.find((frame) => {
      try {
        const url = new URL(frame.src, window.location.origin);
        return url.searchParams.get("frameEditorId") === this.containerId;
      } catch {
        return false;
      }
    });
    if (matchedFrame) {
      return matchedFrame;
    }
    if (this.containerId === ONLYOFFICE_ID) {
      return frames[0];
    }
    return document.querySelector(
      `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"]`
    )?.querySelector('iframe[name="frameEditor"]');
  }
  installProxiesOnWindow(win) {
    installOnlyOfficeProxies(win, this.server, this.createScopedIo());
  }
  /**
   * 劫持 iframe 内 XHR/fetch/io，将协作与 downloadAs 请求路由到 mock EditorServer。
   * 必须在 downloadAs 前安装，否则 export 无法收到 /downloadas/ 分片。
   */
  installIframeProxies() {
    const iframe = this.getEditorFrameElement();
    if (!iframe) {
      throw new Error("Iframe not loaded");
    }
    if (!canAccessIframeWindow(iframe)) {
      return;
    }
    const win = iframe.contentWindow;
    const iframeDoc = iframe.contentDocument;
    if (!iframeDoc || !win) {
      throw new Error("Iframe not loaded");
    }
    if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
      return;
    }
    this.installProxiesOnWindow(win);
    installReporterWindowHook(win, (target) => {
      this.installProxiesOnWindow(target);
    });
    this.installSaveShortcutBlocker();
  }
  getEditorFrameWindow() {
    const iframe = this.getEditorFrameElement();
    if (!iframe || !canAccessIframeWindow(iframe)) {
      return void 0;
    }
    return iframe.contentWindow;
  }
  getSdkApi() {
    return this.getEditorFrameWindow()?.Asc?.editor;
  }
  requireSdkApi() {
    const api = this.getSdkApi();
    if (!api) {
      throw new Error("OnlyOffice SDK API is not ready");
    }
    return api;
  }
  installCommentResolveCleanup() {
    const api = this.getSdkApi();
    if (!api || api.__ONLYOFFICE_RESOLVE_PATCHED__) {
      return;
    }
    api.__ONLYOFFICE_RESOLVE_PATCHED__ = true;
    const removeResolvedComment = (id) => {
      window.setTimeout(() => {
        api.asc_removeComment?.(String(id));
        this.comments.delete(String(id));
      }, 0);
    };
    const originalSyncChange = api.sync_ChangeCommentData?.bind(api);
    if (originalSyncChange) {
      api.sync_ChangeCommentData = (id, data, ...args) => {
        const result = originalSyncChange(id, data, ...args);
        if (isResolvedComment(data)) {
          removeResolvedComment(id);
        }
        return result;
      };
    }
    api.asc_registerCallback?.("asc_onChangeCommentData", (id, data) => {
      if (!isResolvedComment(data)) {
        return;
      }
      removeResolvedComment(id);
    });
  }
  getDocumentPermissions(editing) {
    const doc = this.server.getDocument();
    return {
      edit: editing && doc.fileType !== "pdf",
      chat: false,
      rename: editing,
      protect: editing,
      // 允许接受/拒绝文档内已有修订；不自动进入「修订」录制模式
      review: true,
      print: false
    };
  }
  /** 关闭 autosave 与保存按钮；保存快捷键由 installSaveShortcutBlocker 拦截。 */
  buildEditorCustomization() {
    return {
      uiTheme: this.uiTheme,
      autosave: false,
      layout: {
        header: {
          save: false,
          // Word 头部「编辑 / 审阅 / 查看」切换（PPT/Excel 无此入口）
          editMode: false
        },
        toolbar: {
          file: {
            save: false
          },
          save: false
        }
      },
      review: {
        trackChanges: this.revisionReviewMode,
        // showReviewChanges:true 会在加载时弹出 asc-window review-changes modal-dlg
        showReviewChanges: false,
        ...this.revisionReviewMode ? { reviewDisplay: "markup" } : {}
      },
      features: {
        spellcheck: {
          change: false
        }
      }
      /**
       * DSH patch: upstream sets `logo.image/imageDark` to a
       * `cdn.jsdelivr.net/npm/simple-icons/...` URL, which makes the editor
       * iframe request a third-party host (and shows a broken logo offline).
       * The plugin ships no branding asset and must stay self-contained, so the
       * customization omits `logo` and the editor uses its built-in logo.
       */
    };
  }
  /** 禁用 Ctrl/Cmd+S 与工具栏保存，避免与 export/downloadAs 共用管道冲突。 */
  /** DSH patch: the app's write-back runs from the editor's own save affordances. */
  saveShortcutHandler = null;
  /** DSH patch: notified when the document content becomes dirty (auto-save). */
  changeHandler = null;
  /** DSH patch: register the auto-save trigger for document changes. */
  setChangeHandler(handler) {
    this.changeHandler = handler;
  }
  /** DSH patch: register what Ctrl/Cmd+S inside the editor should do. */
  setSaveShortcutHandler(handler) {
    this.saveShortcutHandler = handler;
  }
  installSaveShortcutBlocker() {
    const win = this.getEditorFrameWindow();
    const doc = win?.document;
    if (!doc || win?.__ONLYOFFICE_SAVE_BLOCKED__) {
      return;
    }
    const blockSaveShortcut = (event) => {
      const key = event.key.toLowerCase();
      if ((event.ctrlKey || event.metaKey) && key === "s") {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        this.saveShortcutHandler?.();
      }
    };
    doc.addEventListener("keydown", blockSaveShortcut, true);
    win.__ONLYOFFICE_SAVE_BLOCKED__ = true;
  }
  /** 文档若自带 w:trackRevisions，OnlyOffice 默认会跟进入修订模式；接入层强制关闭录制。 */
  applyDefaultReviewSettings() {
    this.trackRevisions = false;
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SET_TRACK,
        {
          enabled: false
        }
      ).catch(() => {
      });
      return;
    }
    const api = this.getSdkApi();
    api?.asc_SetGlobalTrackRevisions?.(false);
  }
  mergeCommentItems(items) {
    for (const item of items) {
      if (isResolvedComment(item.Data)) {
        this.comments.delete(item.Id);
        continue;
      }
      this.comments.set(item.Id, item.Data);
    }
  }
  fetchCommentsFromSdk() {
    const raw = this.getSdkApi()?.pluginMethod_GetAllComments?.();
    if (!Array.isArray(raw)) {
      return [];
    }
    return raw.map((item, index) => {
      const source = item && typeof item === "object" ? item : {};
      const id = String(source.Id ?? source.id ?? `comment-${index}`);
      const data = source.Data ?? source.data ?? source;
      return { Id: id, Data: data };
    }).filter((item) => !isResolvedComment(item.Data));
  }
  refreshCommentsFromSdk() {
    if (isOnlyOfficeCdnMode()) {
      void this.refreshCrossOriginComments().catch(() => {
      });
      return;
    }
    this.mergeCommentItems(this.fetchCommentsFromSdk());
  }
  normalizeCrossOriginComments(value) {
    if (!Array.isArray(value)) {
      return [];
    }
    return value.map((item, index) => {
      const source = item && typeof item === "object" ? item : {};
      const data = source.Data && typeof source.Data === "object" ? source.Data : {};
      return {
        Id: String(source.Id ?? `comment-${index}`),
        Data: data
      };
    });
  }
  refreshCrossOriginComments() {
    if (this.crossOriginCommentRefreshPromise) {
      return this.crossOriginCommentRefreshPromise;
    }
    this.crossOriginCommentRefreshPromise = this.callCrossOriginComment(
      CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_LIST,
      {}
    ).then((items) => {
      const normalized = this.normalizeCrossOriginComments(items);
      this.comments.clear();
      this.mergeCommentItems(normalized);
      return Array.from(this.comments.entries()).map(([Id, Data]) => ({
        Id,
        Data
      }));
    }).finally(() => {
      this.crossOriginCommentRefreshPromise = null;
    });
    return this.crossOriginCommentRefreshPromise;
  }
  refreshRevisionsFromSdk(options) {
    if (isOnlyOfficeCdnMode()) {
      void this.refreshCrossOriginRevisions(options).catch(() => {
      });
      return;
    }
    if (this.refreshingRevisions) return;
    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      this.revisions = [];
      return;
    }
    this.refreshingRevisions = true;
    try {
      this.revisions = collectRevisionItems(
        api,
        frameWin,
        options
      );
    } finally {
      this.refreshingRevisions = false;
    }
  }
  applyRevisionsFromSdkStack(stack) {
    if (isOnlyOfficeCdnMode()) {
      this.revisions = this.normalizeCrossOriginRevisions(stack);
      return;
    }
    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) return;
    this.revisions = resolveRevisionShowChanges(
      stack,
      api,
      frameWin
    );
  }
  syncRevisionsAfterMutation() {
    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
  }
  applyAllRevisionChanges(mode) {
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        mode === "accept" ? CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT_ALL : CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT_ALL
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
      return;
    }
    const api = this.requireSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!frameWin) {
      return;
    }
    const applyBulk = mode === "accept" ? () => api.asc_AcceptChanges?.() : () => api.asc_RejectChanges?.();
    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
    let guard = 0;
    let stagnant = 0;
    let lastCount = this.revisions.length;
    while (this.haveRevisionsChanges() && guard++ < 20) {
      this.refreshRevisionsFromSdk({ forceRefreshStack: true });
      const [first] = this.revisions;
      if (!first) {
        applyBulk();
        this.syncRevisionsAfterMutation();
        continue;
      }
      applyRevisionChange(mode, first, api, frameWin, this.revisions);
      this.syncRevisionsAfterMutation();
      const nextCount = this.revisions.length;
      if (nextCount >= lastCount && this.haveRevisionsChanges()) {
        stagnant += 1;
        if (stagnant >= 3) {
          applyBulk();
          this.syncRevisionsAfterMutation();
          stagnant = 0;
        }
      } else {
        stagnant = 0;
      }
      lastCount = nextCount;
    }
    this.syncRevisionsAfterMutation();
  }
  normalizeCrossOriginRevisions(value) {
    if (!Array.isArray(value)) {
      return [];
    }
    return value.map((item, index) => {
      const source = item && typeof item === "object" ? item : {};
      const data = source.Data && typeof source.Data === "object" ? source.Data : {};
      return {
        Id: String(source.Id ?? `rev-stack-${index}`),
        Index: typeof source.Index === "number" && Number.isFinite(source.Index) ? source.Index : index,
        Data: data,
        Raw: source.Raw ?? {}
      };
    });
  }
  refreshCrossOriginRevisions(options) {
    if (this.crossOriginRevisionRefreshPromise) {
      return this.crossOriginRevisionRefreshPromise;
    }
    this.crossOriginRevisionRefreshPromise = this.callCrossOriginRevision(
      CROSS_ORIGIN_EDITOR_COMMAND.REVISION_LIST,
      {
        forceRefreshStack: !!options?.forceRefreshStack
      }
    ).then((items) => {
      this.revisions = this.normalizeCrossOriginRevisions(items);
      return this.revisions;
    }).finally(() => {
      this.crossOriginRevisionRefreshPromise = null;
    });
    return this.crossOriginRevisionRefreshPromise;
  }
  callCrossOriginRevision(command, payload = {}) {
    return callCrossOriginEditor(this.containerId, command, payload);
  }
  revisionTargetId(revision) {
    return typeof revision === "string" ? revision : revision.Id;
  }
  addDemoRevision(text = `审批修订 ${(/* @__PURE__ */ new Date()).toLocaleTimeString()}`) {
    this.trackRevisions = true;
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ADD_DEMO,
        { text }
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
        return this.revisions;
      });
    }
    const api = this.requireSdkApi();
    api.asc_SetGlobalTrackRevisions?.(true);
    api.asc_SetLocalTrackRevisions?.(true);
    if (api.pluginMethod_InputText) {
      api.pluginMethod_InputText(text);
    } else if (api.pluginMethod_PasteText) {
      api.pluginMethod_PasteText(text);
    } else if (api.asc_AddText) {
      api.asc_AddText(text);
    } else {
      throw new Error("OnlyOffice text insertion API is not available");
    }
    this.syncRevisionsAfterMutation();
    return this.revisions;
  }
  teardownWordContentSync() {
    this.wordContentSyncTeardown?.();
    this.wordContentSyncTeardown = null;
    this.wordContentSyncPromise = null;
  }
  scheduleWordContentSync() {
    window.setTimeout(() => {
      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();
    }, 0);
  }
  ensureWordContentSync() {
    if (this.fileType !== "docx" && getDocumentType(this.fileType) !== "word") {
      return Promise.resolve();
    }
    if (this.wordContentSyncPromise) {
      return this.wordContentSyncPromise;
    }
    this.wordContentSyncPromise = (async () => {
      if (isOnlyOfficeCdnMode()) {
        await Promise.all([
          this.refreshCrossOriginComments(),
          this.refreshCrossOriginRevisions(),
          this.callCrossOriginComment(
            CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
            {}
          ),
          this.callCrossOriginRevision(
            CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE
          )
        ]);
        const unsubscribers2 = [
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                return;
              }
              this.comments.set(commentId, commentData);
            }
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                return;
              }
              this.comments.set(commentId, commentData);
            }
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT,
            ([id]) => {
              this.comments.delete(String(id));
            }
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            ([items]) => {
              this.revisions = this.normalizeCrossOriginRevisions(items);
            }
          )
        ];
        this.wordContentSyncTeardown = () => {
          unsubscribers2.forEach((unsubscribe) => unsubscribe());
        };
        return;
      }
      const api = this.requireSdkApi();
      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();
      const unsubscribers = await Promise.all([
        this.subscribe({
          type: "asc_onAddComment",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }
            this.comments.set(commentId, commentData);
          }
        }),
        this.subscribe({
          type: "asc_onChangeCommentData",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }
            this.comments.set(commentId, commentData);
          }
        }),
        this.subscribe({
          type: "asc_onRemoveComment",
          fn: (id) => {
            this.comments.delete(String(id));
          }
        }),
        this.subscribe({
          type: CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
          fn: (stack) => {
            this.applyRevisionsFromSdkStack(stack);
          }
        })
      ]);
      this.wordContentSyncTeardown = () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe?.());
      };
      this.scheduleWordContentSync();
      api.asc_Recalculate?.();
    })().catch(() => {
      this.teardownWordContentSync();
    });
    return this.wordContentSyncPromise;
  }
  getRestrictionSdkApi() {
    return this.getSdkApi();
  }
  getShellMainController() {
    const win = this.getEditorFrameWindow();
    return win?.PE?.getController?.("Main") ?? win?.getApplication?.()?.getController?.("Main") ?? win?.DE?.getController?.("Main");
  }
  getWordHeaderView() {
    const win = this.getEditorFrameWindow();
    return win?.DE?.getController?.("Viewport")?.getView?.(
      "Common.Views.Header"
    );
  }
  /** 隐藏 Word 头部「编辑 / 审阅 / 查看」切换（customization + 运行时兜底）。 */
  hideWordDocModeSwitcher() {
    if (getDocumentType(this.fileType) !== "word" /* Word */) {
      return;
    }
    const header = this.getWordHeaderView();
    header?.btnDocMode?.setVisible?.(false);
    header?.btnPDFMode?.setVisible?.(false);
    const hedset = this.getEditorFrameWindow()?.document?.querySelector(
      '[data-layout-name="header-editMode"]'
    );
    if (hedset instanceof HTMLElement) {
      hedset.style.display = "none";
    }
  }
  scheduleWordDocModeHide() {
    this.hideWordDocModeSwitcher();
    window.setTimeout(() => this.hideWordDocModeSwitcher(), 0);
  }
  /** 同步 web-apps 工具栏/侧栏的 editing:disable（viewMode 与只读一致）。 */
  syncShellEditingDisable(disabled, documentType = getDocumentType(this.fileType)) {
    const nc = this.getEditorFrameWindow()?.Common?.NotificationCenter;
    if (!nc?.trigger) {
      return;
    }
    if (documentType === "slide" /* Slide */) {
      nc.trigger("editing:disable", disabled, {
        viewMode: disabled,
        allowSignature: !disabled,
        rightMenu: { clear: false, disable: true },
        statusBar: true,
        leftMenu: { disable: disabled, previewMode: disabled },
        fileMenu: false,
        comments: { disable: false, previewMode: disabled },
        chat: false,
        review: true,
        viewport: disabled,
        documentHolder: { clear: disabled, disable: true },
        toolbar: true,
        header: { search: false },
        shortcuts: disabled ? false : void 0
      });
      return;
    }
    if (documentType === "word" /* Word */) {
      if (disabled) {
        nc.trigger("editing:disable", true, {
          viewMode: true,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: false,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: true, disable: true },
          statusBar: true,
          leftMenu: { disable: true, previewMode: true },
          fileMenu: { protect: true, history: false },
          navigation: { disable: false, previewMode: true },
          comments: { disable: false, previewMode: true },
          chat: false,
          review: true,
          viewport: true,
          documentHolder: { clear: true, disable: true },
          toolbar: true,
          protect: true,
          header: { search: false, startfill: false },
          shortcuts: false
        });
      } else {
        nc.trigger("editing:disable", false, {
          viewMode: false,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: true,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: false, disable: true },
          statusBar: true,
          leftMenu: { disable: false, previewMode: false },
          fileMenu: false,
          navigation: { disable: false, previewMode: false },
          comments: { disable: false, previewMode: false },
          chat: false,
          review: true,
          viewport: false,
          documentHolder: { clear: false, disable: true },
          toolbar: true,
          protect: true
        });
      }
      this.scheduleWordDocModeHide();
      return;
    }
    nc.trigger("editing:disable", disabled, {
      viewMode: disabled,
      reviewMode: false,
      fillFormMode: false,
      viewDocMode: false,
      allowMerge: true,
      allowSignature: false,
      allowProtect: false,
      rightMenu: { clear: false, disable: true },
      statusBar: true,
      leftMenu: { disable: false, previewMode: disabled },
      fileMenu: false,
      navigation: { disable: false, previewMode: disabled },
      comments: { disable: false, previewMode: disabled },
      chat: false,
      review: true,
      viewport: false,
      documentHolder: { clear: false, disable: true },
      toolbar: true
    });
  }
  /** PPT 在通用只读逻辑之上追加：锁定 Main 控制器 + 禁用幻灯片侧栏。 */
  syncSlideReadOnlyExtras(locked) {
    if (getDocumentType(this.fileType) !== "slide" /* Slide */ || !locked) {
      return;
    }
    this.lockShellEditMode();
  }
  /**
   * processRightsChange(true) 在 OnlyOffice 内无效果；false 会 asc_coAuthoringDisconnect 且 mode.isEdit=false。
   * 本地/mock 场景用 asc_setRestriction + 外壳 UI 同步，避免切回编辑仍停留在只读。
   */
  restoreShellEditMode() {
    const main = this.getShellMainController();
    if (main?.mode) {
      main.mode.isEdit = true;
      main.mode.canEdit = true;
    }
    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(true);
  }
  lockShellEditMode() {
    const main = this.getShellMainController();
    if (main?.mode) {
      main.mode.isEdit = false;
      main.mode.canEdit = false;
    }
    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(false);
  }
  /** 兜底：拦截 SDK 层新增/复制幻灯片（toolbar 锁定之外）。 */
  installSlideStructureEditBlocker() {
    if (getDocumentType(this.fileType) !== "slide" /* Slide */) {
      return;
    }
    const patchApi = (api) => {
      if (!api || api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__) {
        return;
      }
      const guard = (fn) => {
        if (!fn) {
          return fn;
        }
        const bound = fn.bind(api);
        return (...args) => {
          if (this.readOnly) {
            return void 0;
          }
          return bound(...args);
        };
      };
      api.AddSlide = guard(api.AddSlide);
      api.DublicateSlide = guard(api.DublicateSlide);
      api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__ = true;
    };
    patchApi(
      this.getSdkApi()
    );
    patchApi(
      this.getShellMainController()?.api
    );
  }
  /** downloadAs → /downloadas/ → 更新 fsMap 中的 Editor.bin。 */
  async captureDocumentSnapshot() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }
    return await this.server.captureCurrentDocument(() => {
      this.installIframeProxies();
      this.editor?.downloadAs("bin");
    });
  }
  /**
   * 只读模式下 downloadAs 可能被 SDK 拦截；导出前临时恢复编辑权再抓取。
   */
  async captureDocumentSnapshotAllowingReadOnly() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }
    const locked = this.readOnly;
    if (locked) {
      this.syncEditingRights(true);
    }
    try {
      return await this.captureDocumentSnapshot();
    } finally {
      if (locked) {
        this.syncEditingRights(false);
        this.syncSlideReadOnlyExtras(true);
      }
    }
  }
  async captureDocumentIfDirty() {
    if (!this.editor || this.readOnly || !this.dirty) {
      return;
    }
    await this.captureDocumentSnapshot();
    this.dirty = false;
  }
  destroyDocEditorInstance() {
    this.disconnectConnector();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.comments.clear();
    this.revisions = [];
    this.teardownWordContentSync();
  }
  /**
   * 创建 Developer Edition Connector。
   * Connector 运行在父页面，借助 DocsAPI 与编辑器 iframe 通信，因此可用于 CDN 跨域场景。
   */
  createConnector(options) {
    if (!this.editor) {
      throw new Error("OnlyOffice editor is not ready");
    }
    if (this.connector) {
      if (options?.autoconnect !== false && !this.connector.isConnected) {
        this.connector.connect();
      }
      return this.connector;
    }
    const iframe = this.getEditorFrameElement();
    if (!iframe?.contentWindow) {
      throw new Error("OnlyOffice editor iframe is not ready");
    }
    const connector = this.editor.createConnector({
      ...options,
      autoconnect: false
    });
    connector.guid = this.containerId;
    const iframeUrl = new URL(iframe.src, window.location.href);
    const frameEditorId = iframeUrl.searchParams.get("frameEditorId") ?? "iframeEditor";
    connector.sendMessage = (data) => {
      iframe.contentWindow?.postMessage(
        JSON.stringify({
          frameEditorId,
          type: "onExternalPluginMessage",
          subType: "connector",
          data: { ...data, guid: connector.guid }
        }),
        iframeUrl.origin
      );
    };
    if (options?.autoconnect !== false) {
      connector.connect();
    }
    this.connector = connector;
    return this.connector;
  }
  disconnectConnector() {
    if (this.connector?.isConnected) {
      this.connector.disconnect();
    }
    this.connector = null;
  }
  /**
   * 初始只读与运行时切只读走同一套 asc_setRestriction。
   * 挂载阶段若 permissions.edit=false，xlsx 等会在打开时样式/格式异常；
   * 因此挂载时保持完整编辑权限，documentReady 后再施加只读限制。
   * CDN 跨域模式通过 iframe 内 bridge 执行同样的 SDK/UI 同步，避免跨域访问 window.Asc。
   */
  applyInitialReadOnlyState(documentType) {
    if (isOnlyOfficeCdnMode()) {
      this.syncCrossOriginReadOnly(this.readOnly);
      return;
    }
    this.installSlideStructureEditBlocker();
    this.syncEditingRights(false);
    if (documentType === "slide" /* Slide */) {
      window.setTimeout(() => {
        if (this.readOnly) {
          this.syncShellEditingDisable(true, documentType);
        }
      }, 0);
    }
    if (documentType === "cell" /* Cell */) {
      this.getSdkApi()?.asc_Recalculate?.();
    }
  }
  /** 语言写在 iframe URL 的 lang 参数里，运行时 refreshFile 不会更新界面语言。 */
  mountDocEditor() {
    const doc = this.server.getDocument();
    const user = this.server.getUser();
    const documentType = getDocumentType(doc.fileType);
    this.server.setClient({
      buildVersion: window.DocsAPI.DocEditor.version()
    });
    this.editor = new window.DocsAPI.DocEditor(this.containerId, {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(true)
      },
      documentType,
      editorConfig: {
        lang: this.editorLang,
        coEditing: {
          mode: "fast",
          change: false
        },
        user: {
          ...user
        },
        customization: this.buildEditorCustomization()
      },
      events: {
        onAppReady: () => {
          this.installIframeProxies();
        },
        onDocumentReady: () => {
          this.installSaveShortcutBlocker();
          this.installCommentResolveCleanup();
          this.installSlideStructureEditBlocker();
          this.applyDefaultReviewSettings();
          void this.ensureWordContentSync();
          onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, {
            fileName: doc.title,
            fileType: doc.fileType,
            instanceId: this.instanceId
          });
          if (this.readOnly) {
            this.applyInitialReadOnlyState(documentType);
          } else if (documentType === "word" /* Word */) {
            this.scheduleWordDocModeHide();
          }
        },
        onDocumentStateChange: (event) => {
          if (event.data) {
            this.dirty = true;
            this.changeHandler?.();
          }
        },
        // DocsAPI 仅在注册此回调时暴露「文件 → 重命名」入口。
        // 回调由 OnlyOffice 的内部 Gateway 发出；标题显示已由 iframe 更新，
        // 使用 iframe 内 asc_wopi_renameFile 发起 RPC，由内存服务按 WOPI 协议回包。
        // onRequestRename: (event: { data?: unknown }) => {
        //   const fileName =
        //     typeof event.data === "string" ? event.data : "";
        //   void this.renameDocument(fileName).catch((error) => {
        //     this.logger.error("operation", "OnlyOffice document rename failed", {
        //       fileName,
        //       instanceId: this.instanceId,
        //       containerId: this.containerId,
        //       error: error instanceof Error ? error.message : String(error),
        //     });
        //   });
        // },
        // DSH patch: the editor's own save control (Quick Access floppy /
        // File ▸ Save) must reach the app, which writes the bytes back to the
        // workspace file. `notifyUserSave` broadcasts SAVE_DOCUMENT + ONSAVE
        // for listeners; `onDownloadAs` below still serves export().
        onSave: () => {
          this.notifyUserSave();
        },
        onDownloadAs: () => {
        }
      },
      type: "desktop",
      width: "100%",
      height: "100%"
    });
  }
  buildRefreshFileConfig(editing) {
    const doc = this.server.getDocument();
    return {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(editing)
      },
      documentType: getDocumentType(doc.fileType),
      editorConfig: {
        mode: editing ? "edit" : "view",
        lang: this.editorLang
      },
      type: "desktop",
      width: "100%",
      height: "100%"
    };
  }
  /** 就地切换编辑权限（主路径 asc_setRestriction；PPT 只读额外 syncSlideReadOnlyExtras）。 */
  syncEditingRights(editing) {
    if (!this.editor) {
      return;
    }
    const documentType = getDocumentType(this.fileType);
    if (isOnlyOfficeCdnMode()) {
      this.syncCrossOriginReadOnly(!editing);
      return;
    }
    const sdk = this.getRestrictionSdkApi();
    if (sdk?.asc_setRestriction) {
      if (editing) {
        sdk.asc_removeRestriction?.(ASC_RESTRICTION_VIEW);
        sdk.asc_setRestriction(ASC_RESTRICTION_NONE);
        this.restoreShellEditMode();
      } else {
        sdk.asc_setRestriction(ASC_RESTRICTION_VIEW);
        this.syncSlideReadOnlyExtras(true);
      }
      this.syncShellEditingDisable(!editing, documentType);
      return;
    }
    if (editing) {
      this.restoreShellEditMode();
      this.syncShellEditingDisable(false, documentType);
      this.editor.refreshFile?.(this.buildRefreshFileConfig(true));
    } else {
      this.syncSlideReadOnlyExtras(true);
      if (documentType !== "slide" /* Slide */) {
        this.editor.denyEditingRights?.("");
      }
      this.syncShellEditingDisable(true, documentType);
    }
  }
  createExportData(snapshot) {
    const binData = snapshot.binData;
    if (!binData) {
      throw new Error("No OnlyOffice document data is available to export");
    }
    return {
      fileName: snapshot.fileName || this.fileName,
      fileType: snapshot.fileType || this.fileType,
      binData,
      instanceId: this.instanceId,
      media: {
        ...snapshot.media,
        ...this.media
      },
      themes: snapshot.themes
    };
  }
  userSaveTimer = null;
  pendingUserSaveSnapshot = null;
  /** 用户保存：更新快照并广播 SAVE_DOCUMENT + ONSAVE（同 tick 内合并重复回调）。 */
  notifyUserSave(snapshot) {
    if (snapshot) {
      this.pendingUserSaveSnapshot = snapshot;
    }
    if (this.userSaveTimer !== null) {
      window.clearTimeout(this.userSaveTimer);
    }
    this.userSaveTimer = window.setTimeout(() => {
      this.userSaveTimer = null;
      this.dirty = false;
      const snap = this.pendingUserSaveSnapshot ?? this.server.getDocumentSnapshot();
      this.pendingUserSaveSnapshot = null;
      const data = this.createExportData(snap);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.ONSAVE, {
        fileName: data.fileName,
        instanceId: this.instanceId
      });
    }, 0);
  }
  /**
   * DSH patch: re-open a changed document INSIDE the live editor.
   *
   * Switching documents normally destroys and recreates the whole DocEditor
   * (mountDocEditor). For an external file change that is a visible full
   * reload; DocsAPI's `refreshFile` re-opens just the document in the existing
   * instance, so the editor chrome (toolbar, layout) stays put.
   *
   * @param file new document bytes
   * @param fileType optional macro type override
   * @returns whether the refresh was issued
   */
  async refreshDocument(file, fileType) {
    if (!this.editor) return false;
    const type = getFileType(file.name, fileType);
    this.fileName = file.name;
    this.fileType = type;
    await this.server.open(file, { fileName: file.name, fileType: type });
    this.editor.refreshFile?.(this.buildRefreshFileConfig(true));
    return true;
  }
  isLoadSessionActive(containerId, loadSession) {
    return loadSession === void 0 || editorManagerFactory.isLoadSessionActive(containerId, loadSession);
  }
  async create(options) {
    const containerId = options.containerId || this.containerId || ONLYOFFICE_ID;
    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }
    this.destroy();
    this.readOnly = !!options.readOnly;
    this.revisionReviewMode = !!options.revisionReview;
    this.server.setOfficeXmlEventConfig(options.officeXmlEvent);
    if (options.user) {
      this.server.setUser(options.user);
    }
    this.containerId = containerId;
    const fileType = getFileType(options.fileName, options.fileType);
    this.fileName = options.fileName;
    this.fileType = fileType;
    this.media = {};
    this.comments.clear();
    this.revisions = [];
    this.clearOfficeXmlSizeLimitOverlay();
    this.teardownWordContentSync();
    if (options.isNew) {
      this.server.openNew(fileType, options.fileName);
    } else if (options.file) {
      await this.server.open(options.file, {
        fileName: options.fileName,
        fileType
      });
    } else if (options.url) {
      await this.server.openUrl(options.url, {
        fileName: options.fileName,
        fileType,
        loader: options.loader
      });
    } else {
      throw new Error("OnlyOffice requires a file, url, or new document type");
    }
    if (this.officeXmlSizeLimitPayload) {
      this.renderOfficeXmlSizeLimitOverlay();
      return this;
    }
    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }
    await initializeOnlyOffice();
    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }
    this.editorLang = options.lang || getOnlyOfficeLang();
    this.uiTheme = options.theme || "theme-white";
    this.syncEditorBridge();
    this.mountDocEditor();
    this.renderOfficeXmlSizeLimitOverlay();
    return this;
  }
  exists() {
    return !!this.editor;
  }
  /** 导出链路：downloadAs("bin") → server.resolvePendingExport → SAVE_DOCUMENT 事件。 */
  async export() {
    let snapshot;
    if (this.editor && (!this.readOnly || this.dirty)) {
      snapshot = await this.captureDocumentSnapshotAllowingReadOnly();
      this.dirty = false;
    } else {
      snapshot = this.server.getDocumentSnapshot();
    }
    const data = this.createExportData(snapshot);
    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);
    return data;
  }
  getUser() {
    return this.server.getUser();
  }
  setUser(user) {
    this.server.setUser(user);
    this.editor?.setUsers?.([{ id: user.id, name: user.name }]);
  }
  /** 就地切换只读；切到只读前先 downloadAs 落盘，避免后续导出仍是打开时的 Editor.bin。 */
  async setReadOnly(readOnly) {
    if (this.readOnly === readOnly) {
      return;
    }
    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true
    });
    try {
      if (readOnly && this.editor) {
        await this.captureDocumentSnapshot();
        this.dirty = false;
      }
      this.readOnly = readOnly;
      if (isOnlyOfficeCdnMode()) {
        this.syncCrossOriginReadOnly(readOnly);
        return;
      }
      this.installSlideStructureEditBlocker();
      this.syncEditingRights(!readOnly);
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false
      });
    }
  }
  getReadOnly() {
    return this.readOnly;
  }
  async setLanguage(lang) {
    if (this.editorLang === lang) {
      return;
    }
    this.editorLang = lang;
    if (!this.editor) {
      return;
    }
    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true
    });
    try {
      await this.captureDocumentIfDirty();
      this.destroyDocEditorInstance();
      this.mountDocEditor();
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false
      });
    }
  }
  getTheme() {
    return this.uiTheme;
  }
  /** uiTheme 写在 iframe URL 参数里，运行时需 remount 才能生效。 */
  async setTheme(theme) {
    if (this.uiTheme === theme) {
      return;
    }
    this.uiTheme = theme;
    if (!this.editor) {
      return;
    }
    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true
    });
    try {
      await this.captureDocumentIfDirty();
      this.destroyDocEditorInstance();
      this.mountDocEditor();
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false
      });
    }
  }
  getInstanceId() {
    return this.instanceId;
  }
  isOfficeXmlSizeLimitExceeded() {
    return !!this.officeXmlSizeLimitPayload;
  }
  getContainerId() {
    return this.containerId;
  }
  getLogger() {
    return this.logger;
  }
  printLogs() {
    this.logger.print();
  }
  getFileName() {
    return this.fileName;
  }
  /**
   * @description 通过 iframe 内 asc_wopi_renameFile 重命名当前实例。
  * SDK 经 socket rpc 收到内存 WOPI 回包后才更新本地标题，因此返回 Promise。
   */
  renameDocument(fileName) {
    const requestedName = typeof fileName === "string" ? fileName.trim() : "";
    if (!requestedName) {
      return Promise.reject(
        new Error("OnlyOffice document name cannot be empty")
      );
    }
    if (this.pendingRename) {
      return Promise.reject(new Error("OnlyOffice document rename is pending"));
    }
    return new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        if (this.pendingRename?.timer !== timer) {
          return;
        }
        this.pendingRename = null;
        reject(new Error("OnlyOffice document rename timed out"));
      }, 5e3);
      this.pendingRename = { resolve, reject, timer };
      try {
        if (isOnlyOfficeCdnMode()) {
          void callCrossOriginEditor(
            this.containerId,
            CROSS_ORIGIN_EDITOR_COMMAND.DOCUMENT_RENAME,
            { fileName: requestedName }
          ).then(() => {
            if (this.pendingRename) {
              this.server.rename(requestedName);
            }
          }).catch((error) => this.rejectPendingRename(error));
          return;
        }
        const api = this.requireSdkApi();
        if (!api.asc_wopi_renameFile) {
          throw new Error("OnlyOffice WOPI rename API is not available");
        }
        api.asc_wopi_renameFile(requestedName);
      } catch (error) {
        this.rejectPendingRename(error);
      }
    });
  }
  syncRenamedDocument(renamedFileName) {
    this.fileName = renamedFileName;
    const document2 = getDocumentObj(this.containerId);
    setDocumentObj(
      {
        ...document2,
        fileName: renamedFileName
      },
      this.containerId
    );
    const pendingRename = this.pendingRename;
    if (pendingRename) {
      window.clearTimeout(pendingRename.timer);
      this.pendingRename = null;
      pendingRename.resolve(renamedFileName);
    }
    this.logger.operation("OnlyOffice document renamed through WOPI RPC", {
      fileName: renamedFileName,
      instanceId: this.instanceId,
      containerId: this.containerId
    });
  }
  rejectPendingRename(error) {
    const pendingRename = this.pendingRename;
    if (!pendingRename) {
      return;
    }
    window.clearTimeout(pendingRename.timer);
    this.pendingRename = null;
    pendingRename.reject(
      error instanceof Error ? error : new Error(String(error))
    );
  }
  getContainerParentSelector() {
    return `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"], ${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}`;
  }
  getContainerStyle() {
    return ONLYOFFICE_CONTAINER_CONFIG.STYLE;
  }
  updateMedia(key, data) {
    this.media[key] = data;
  }
  getMedia() {
    return { ...this.media };
  }
  isDirty() {
    return this.dirty;
  }
  async subscribe({
    type,
    fn
  }) {
    if (isOnlyOfficeCdnMode()) {
      if (type === CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT || type === CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT || type === CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT) {
        await this.callCrossOriginComment(
          CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
          {}
        );
        return subscribeCrossOriginEditorEvent(
          this.containerId,
          type,
          (args) => fn(...args)
        );
      }
      if (type === CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE || type === CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE) {
        await this.callCrossOriginRevision(
          CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE
        );
        return subscribeCrossOriginEditorEvent(
          this.containerId,
          type,
          (args) => fn(...args)
        );
      }
      if (type === CROSS_ORIGIN_EDITOR_EVENT.DOCUMENT_MODIFIED_CHANGED) {
        await callCrossOriginEditor(
          this.containerId,
          CROSS_ORIGIN_EDITOR_COMMAND.EDITOR_SUBSCRIBE,
          { event: type }
        );
        return subscribeCrossOriginEditorEvent(
          this.containerId,
          type,
          (args) => fn(...args)
        );
      }
      throw new Error(
        `OnlyOffice cross-origin callback is not supported: ${type}`
      );
    }
    const api = this.requireSdkApi();
    if (!api.asc_registerCallback || !api.asc_unregisterCallback) {
      throw new Error("OnlyOffice callback subscription is not supported");
    }
    api.asc_registerCallback(type, fn);
    return () => {
      api.asc_unregisterCallback?.(type, fn);
    };
  }
  async getAllComments() {
    if (isOnlyOfficeCdnMode()) {
      return this.refreshCrossOriginComments();
    }
    this.refreshCommentsFromSdk();
    return Array.from(this.comments.entries()).map(([Id, Data]) => ({
      Id,
      Data
    }));
  }
  refreshComments() {
    return this.getAllComments();
  }
  createSdkCommentPayload(data) {
    const asc = this.getEditorFrameWindow()?.Asc;
    const CommentDataWord = asc?.asc_CCommentDataWord;
    if (!CommentDataWord) {
      return data;
    }
    const comment = new CommentDataWord(null);
    const payload = toPluginCommentPayload(data);
    if (payload.Text != null) {
      comment.asc_putText?.(String(payload.Text));
    }
    if (payload.UserName != null) {
      comment.asc_putUserName?.(String(payload.UserName));
    }
    if (payload.Time != null) {
      comment.asc_putTime?.(String(payload.Time));
    }
    if (payload.QuoteText != null) {
      comment.asc_putQuoteText?.(String(payload.QuoteText));
    }
    if (typeof payload.Solved === "boolean") {
      comment.asc_putSolved?.(payload.Solved);
    }
    if (payload.UserData != null) {
      comment.asc_putUserData?.(String(payload.UserData));
    }
    return comment;
  }
  async callCrossOriginComment(command, payload) {
    return callCrossOriginEditor(this.containerId, command, payload);
  }
  addComment(input) {
    if (isOnlyOfficeCdnMode()) {
      const data2 = toPluginCommentPayload(normalizeCommentInput(input));
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_ADD,
        { data: data2 }
      ).then((id2) => {
        if (id2) {
          this.comments.set(String(id2), data2);
        }
        return id2 ? String(id2) : "";
      });
    }
    const api = this.requireSdkApi();
    const data = toPluginCommentPayload(normalizeCommentInput(input));
    const id = api.pluginMethod_AddComment?.(data) ?? api.asc_addComment?.(this.createSdkCommentPayload(data));
    if (id) {
      this.comments.set(String(id), data);
    }
    return id ? String(id) : "";
  }
  updateComment(id, data) {
    if (isResolvedComment(data)) {
      return this.removeComment(id);
    }
    if (isOnlyOfficeCdnMode()) {
      const payload2 = toPluginCommentPayload(data);
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_UPDATE,
        {
          id,
          data: payload2
        }
      ).then(() => {
        this.comments.set(id, payload2);
      });
    }
    const api = this.requireSdkApi();
    const payload = toPluginCommentPayload(data);
    if (typeof api.pluginMethod_ChangeComment === "function") {
      api.pluginMethod_ChangeComment(id, payload);
    } else {
      api.asc_changeComment?.(
        id,
        this.createSdkCommentPayload(payload)
      );
    }
    this.comments.set(id, payload);
  }
  removeComment(id) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_REMOVE,
        { id }
      ).then(() => {
        this.comments.delete(id);
      });
    }
    this.requireSdkApi().asc_removeComment?.(id);
    this.comments.delete(id);
  }
  goToComment(id, { showBalloon = false } = {}) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_GO_TO,
        { id, showBalloon }
      );
    }
    const api = this.requireSdkApi();
    api.asc_selectComment?.(id);
    if (showBalloon) {
      api.asc_showComment?.(id);
    }
  }
  async registerCommentCallbacks(handlers) {
    if (isOnlyOfficeCdnMode()) {
      const unsubscribers2 = [];
      await this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
        {}
      );
      if (handlers.onAdd) {
        unsubscribers2.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data;
              this.comments.set(commentId, commentData);
              handlers.onAdd?.(commentId, commentData);
            }
          )
        );
      }
      if (handlers.onChange) {
        unsubscribers2.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                handlers.onRemove?.(commentId);
                return;
              }
              this.comments.set(commentId, commentData);
              handlers.onChange?.(commentId, commentData);
            }
          )
        );
      }
      if (handlers.onRemove) {
        unsubscribers2.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT,
            ([id]) => {
              const commentId = String(id);
              this.comments.delete(commentId);
              handlers.onRemove?.(commentId);
            }
          )
        );
      }
      return () => {
        unsubscribers2.forEach((unsubscribe) => unsubscribe());
      };
    }
    const unsubscribers = await Promise.all([
      handlers.onAdd ? this.subscribe({
        type: "asc_onAddComment",
        fn: (id, data) => {
          const commentId = String(id);
          const commentData = data;
          this.comments.set(commentId, commentData);
          handlers.onAdd?.(commentId, commentData);
        }
      }) : void 0,
      handlers.onChange ? this.subscribe({
        type: "asc_onChangeCommentData",
        fn: (id, data) => {
          const commentId = String(id);
          const commentData = data;
          if (isResolvedComment(commentData)) {
            window.setTimeout(() => {
              this.removeComment(commentId);
              handlers.onRemove?.(commentId);
            }, 0);
            return;
          }
          this.comments.set(commentId, commentData);
          handlers.onChange?.(commentId, commentData);
        }
      }) : void 0,
      handlers.onRemove ? this.subscribe({
        type: "asc_onRemoveComment",
        fn: (id) => {
          const commentId = String(id);
          this.comments.delete(commentId);
          handlers.onRemove?.(commentId);
        }
      }) : void 0
    ]);
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }
  setTrackRevisions(enabled) {
    this.trackRevisions = enabled;
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SET_TRACK,
        { enabled }
      );
    }
    const api = this.requireSdkApi();
    api.asc_SetGlobalTrackRevisions?.(enabled);
    api.asc_SetLocalTrackRevisions?.(enabled);
  }
  /** 修订审阅页初始化：开启追踪 + markup 显示模式（不会批量接受/拒绝） */
  prepareRevisionReview() {
    this.trackRevisions = true;
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_PREPARE_REVIEW
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }
    const api = this.requireSdkApi();
    api.asc_SetGlobalTrackRevisions?.(true);
    prepareRevisionReviewDisplay(api, this.getEditorFrameWindow());
  }
  isTrackRevisions() {
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_IS_TRACK
      ).then((enabled) => {
        this.trackRevisions = !!enabled;
      });
      return this.trackRevisions;
    }
    return !!this.getSdkApi()?.asc_GetGlobalTrackRevisions?.();
  }
  haveRevisionsChanges() {
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_HAVE_CHANGES
      ).then((hasChanges) => {
        if (!hasChanges) {
          this.revisions = [];
        } else {
          void this.refreshCrossOriginRevisions();
        }
      });
      return this.revisions.length > 0;
    }
    const api = this.getSdkApi();
    if (typeof api?.asc_HaveRevisionsChanges === "function") {
      if (api.asc_HaveRevisionsChanges(true)) {
        return true;
      }
      if (api.asc_HaveRevisionsChanges()) {
        return true;
      }
    }
    return this.revisions.length > 0;
  }
  async getAllRevisions() {
    if (isOnlyOfficeCdnMode()) {
      return this.refreshCrossOriginRevisions({ forceRefreshStack: true });
    }
    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
    return this.revisions;
  }
  refreshRevisions() {
    return this.getAllRevisions();
  }
  goToNextRevision() {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_NEXT
      );
    }
    this.getSdkApi()?.asc_GetNextRevisionsChange?.();
  }
  goToPrevRevision() {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_PREV
      );
    }
    this.getSdkApi()?.asc_GetPrevRevisionsChange?.();
  }
  goToRevision(id) {
    if (isOnlyOfficeCdnMode()) {
      const cached2 = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_GO_TO,
        {
          id,
          index: cached2?.Index
        }
      );
    }
    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) return;
    let cached = this.revisions.find((entry) => entry.Id === id);
    if (!cached) {
      this.refreshRevisionsFromSdk();
      cached = this.revisions.find((entry) => entry.Id === id);
    }
    goToRevision(
      cached ?? id,
      api,
      frameWin,
      this.revisions
    );
  }
  acceptRevision(revision) {
    if (isOnlyOfficeCdnMode()) {
      const id = this.revisionTargetId(revision);
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT,
        {
          id,
          index: cached?.Index
        }
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }
    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      return;
    }
    if (applyRevisionChange(
      "accept",
      revision,
      api,
      frameWin,
      this.revisions
    )) {
      this.syncRevisionsAfterMutation();
    }
  }
  rejectRevision(revision) {
    if (isOnlyOfficeCdnMode()) {
      const id = this.revisionTargetId(revision);
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT,
        {
          id,
          index: cached?.Index
        }
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }
    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      return;
    }
    if (applyRevisionChange(
      "reject",
      revision,
      api,
      frameWin,
      this.revisions
    )) {
      this.syncRevisionsAfterMutation();
    }
  }
  acceptAllRevisions() {
    this.applyAllRevisionChanges("accept");
  }
  rejectAllRevisions() {
    this.applyAllRevisionChanges("reject");
  }
  acceptRevisionsBySelection(all) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT_SELECTION,
        { all }
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }
    this.requireSdkApi().asc_AcceptChangesBySelection?.(all);
  }
  rejectRevisionsBySelection(all) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT_SELECTION,
        { all }
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }
    this.requireSdkApi().asc_RejectChangesBySelection?.(all);
  }
  async registerRevisionCallbacks(handlers) {
    if (isOnlyOfficeCdnMode()) {
      const unsubscribers2 = [];
      await this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE
      );
      if (handlers.onShowChanges) {
        unsubscribers2.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            ([items]) => {
              this.revisions = this.normalizeCrossOriginRevisions(items);
              handlers.onShowChanges?.(this.revisions);
            }
          )
        );
      }
      if (handlers.onTrackRevisionsChange) {
        unsubscribers2.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
            ([enabled]) => {
              this.trackRevisions = !!enabled;
              handlers.onTrackRevisionsChange?.(!!enabled);
            }
          )
        );
      }
      return () => {
        unsubscribers2.forEach((unsubscribe) => unsubscribe());
      };
    }
    const unsubscribers = await Promise.all([
      handlers.onShowChanges ? this.subscribe({
        type: CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
        fn: (stack) => {
          const api = this.getSdkApi();
          const frameWin = this.getEditorFrameWindow();
          if (!api || !frameWin) return;
          handlers.onShowChanges?.(
            resolveRevisionShowChanges(
              stack,
              api,
              frameWin
            )
          );
        }
      }) : void 0,
      handlers.onTrackRevisionsChange ? this.subscribe({
        type: CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
        fn: (enabled) => handlers.onTrackRevisionsChange?.(!!enabled)
      }) : void 0
    ]);
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }
  destroy() {
    this.rejectPendingRename(new Error("OnlyOffice editor was destroyed"));
    if (this.userSaveTimer !== null) {
      window.clearTimeout(this.userSaveTimer);
      this.userSaveTimer = null;
    }
    this.pendingUserSaveSnapshot = null;
    this.teardownWordContentSync();
    this.crossOriginBridgeTeardown?.();
    this.crossOriginBridgeTeardown = null;
    this.clearOfficeXmlSizeLimitOverlay();
    unregisterScopedIo(this.containerId);
    unregisterCrossOriginBridge(this.containerId);
    this.disconnectConnector();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.dirty = false;
    this.comments.clear();
    this.revisions = [];
    this.server.reset();
  }
};
var EditorManagerFactory = class {
  defaultManager = new EditorManager();
  managers = /* @__PURE__ */ new Map();
  loadSessions = /* @__PURE__ */ new Map();
  beginLoadSession(containerId) {
    const next = (this.loadSessions.get(containerId) ?? 0) + 1;
    this.loadSessions.set(containerId, next);
    return next;
  }
  isLoadSessionActive(containerId, loadSession) {
    return this.loadSessions.get(containerId) === loadSession;
  }
  getDefault() {
    return this.defaultManager;
  }
  create(containerId) {
    const manager = this.managers.get(containerId) || new EditorManager(containerId);
    this.managers.set(containerId, manager);
    return manager;
  }
  get(containerId) {
    return this.managers.get(containerId) || this.create(containerId);
  }
  getAll() {
    return [this.defaultManager, ...this.managers.values()];
  }
  destroy(containerId) {
    const manager = this.managers.get(containerId);
    manager?.destroy();
    this.managers.delete(containerId);
  }
  destroyAll() {
    this.defaultManager.destroy();
    for (const manager of this.managers.values()) {
      manager.destroy();
    }
    this.managers.clear();
  }
};
var editorManagerFactory = new EditorManagerFactory();
var editorManager = editorManagerFactory.getDefault();
if (typeof window !== "undefined") {
  window.editorManagerFactory = editorManagerFactory;
}

// vendor/onlyoffice-web-comp/util/x2t.ts
async function convertBinToDocument(binData, fileName, fileType, media, themes, logger) {
  const targetExt = normalizeX2tExportFileType(fileType);
  const data = new Uint8Array(binData).buffer;
  const { formatFrom, formatTo } = getX2tExportFormats(targetExt);
  const result = await converter.convert(
    {
      data,
      fileFrom: "Editor.bin",
      fileTo: `doc.${targetExt}`,
      formatFrom,
      formatTo,
      media,
      themes
    },
    logger
  );
  if (!result.output) {
    throw new Error("Failed to convert OnlyOffice bin document");
  }
  return {
    fileName: ensureTitleWithExtension(fileName, targetExt),
    data: result.output
  };
}

// vendor/onlyoffice-web-comp/util/download.ts
var OFFICE_MIME_TYPES = {
  [FILE_TYPE.DOCX]: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  [FILE_TYPE.XLSX]: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  [FILE_TYPE.PPTX]: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
};
function downloadBlob(blob, fileName) {
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
}

// vendor/onlyoffice-web-comp/core/onlyoffice-manager.ts
var OnlyOfficeManager = class _OnlyOfficeManager {
  containerId;
  fileType;
  editor;
  readOnly;
  theme;
  officeXmlEvent;
  ready = false;
  constructor(editor, options) {
    this.containerId = options.containerId;
    this.fileType = options.fileType;
    this.editor = editor;
    this.readOnly = options.readOnly ?? false;
    this.theme = options.theme ?? DEFAULT_OFFICE_THEME;
    this.officeXmlEvent = options.officeXmlEvent;
    if (options.user) {
      editor.setUser(options.user);
    }
    if (options.lang) {
      setCurrentLang(options.lang);
    }
  }
  /** 主实例初始化前，可运行时注册 OnlyOffice / x2t 静态资源地址。 */
  static registerStaticResource(options) {
    return registerOnlyOfficeStaticResource(options);
  }
  /** 清空运行时注册地址，恢复默认静态资源地址。 */
  static resetStaticResource() {
    return resetOnlyOfficeStaticResource();
  }
  /** 多实例场景：绑定已有 EditorManager，不自动 open */
  static fromEditor(editor, options) {
    return new _OnlyOfficeManager(editor, options);
  }
  /** 加载 DocsAPI 并打开 defaultFileName（新建或空白） */
  static async create(options) {
    const containerId = options.containerId ?? ONLYOFFICE_ID;
    await initializeOnlyOffice();
    const editor = editorManagerFactory.get(containerId);
    const manager = new _OnlyOfficeManager(editor, { ...options, containerId });
    await manager.openDocument({
      fileName: options.defaultFileName,
      isNew: true,
      readOnly: options.readOnly,
      loadSession: options.loadSession,
      officeXmlEvent: options.officeXmlEvent
    });
    return manager;
  }
  /** 加载 DocsAPI 并直接打开已有 File（先取文件再挂载） */
  static async createWithFile(options, file) {
    const containerId = options.containerId ?? ONLYOFFICE_ID;
    await initializeOnlyOffice();
    const editor = editorManagerFactory.get(containerId);
    const manager = new _OnlyOfficeManager(editor, { ...options, containerId });
    await manager.openDocument({
      fileName: file.name,
      file,
      readOnly: options.readOnly,
      loadSession: options.loadSession,
      officeXmlEvent: options.officeXmlEvent
    });
    return manager;
  }
  /** 打开/切换文档（上传、新建、重开） */
  async openDocument(input) {
    const readOnly = input.readOnly ?? this.readOnly;
    setDocumentObj(
      {
        fileName: input.fileName,
        file: input.file,
        isNew: input.isNew ?? !input.file
      },
      this.containerId
    );
    const { fileName, file } = getDocumentObj(this.containerId);
    await this.editor.create({
      file,
      fileName,
      isNew: !file,
      readOnly,
      user: this.editor.getUser(),
      lang: getOnlyOfficeLang(),
      theme: this.theme,
      containerId: this.containerId,
      editorManager: this.editor,
      loadSession: input.loadSession,
      officeXmlEvent: input.officeXmlEvent ?? this.officeXmlEvent
    });
    this.readOnly = readOnly;
    this.ready = true;
  }
  async openNew(fileName, readOnly) {
    await this.openDocument({ fileName, isNew: true, readOnly });
  }
  async openFile(file, readOnly) {
    await this.openDocument({ fileName: file.name, file, readOnly });
  }
  isReady() {
    return this.ready;
  }
  getReadOnly() {
    return this.readOnly;
  }
  getUser() {
    return this.editor.getUser();
  }
  setUser(user) {
    this.editor.setUser(user);
  }
  async setReadOnly(readOnly) {
    await this.editor.setReadOnly(readOnly);
    this.readOnly = readOnly;
  }
  async toggleReadOnly() {
    return this.setReadOnly(!this.readOnly);
  }
  getLanguage() {
    return getCurrentLang();
  }
  async setLanguage(lang) {
    setCurrentLang(lang);
    await this.editor.setLanguage(lang);
  }
  /** 在中/英之间切换并应用到 iframe */
  async toggleLanguage() {
    const nextLang = getCurrentLang() === ONLYOFFICE_LANG_KEY.ZH ? ONLYOFFICE_LANG_KEY.EN : ONLYOFFICE_LANG_KEY.ZH;
    await this.setLanguage(nextLang);
    return nextLang;
  }
  getTheme() {
    return this.editor.getTheme();
  }
  async setTheme(theme) {
    this.theme = theme;
    await this.editor.setTheme(theme);
  }
  /** 在浅色 / 深色主题之间切换 */
  async toggleTheme() {
    const darkThemes = [
      OFFICE_THEME.DARK,
      OFFICE_THEME.NIGHT,
      OFFICE_THEME.CONTRAST_DARK
    ];
    const nextTheme = darkThemes.includes(this.getTheme()) ? OFFICE_THEME.WHITE : OFFICE_THEME.DARK;
    await this.setTheme(nextTheme);
    return nextTheme;
  }
  getEditor() {
    return this.editor;
  }
  /** 创建 OnlyOffice Developer Edition Connector。编辑器销毁或重开时会自动断开。 */
  createConnector(options) {
    if (!this.ready) {
      throw new Error("OnlyOffice editor is not ready");
    }
    return this.editor.createConnector(options);
  }
  getLogger() {
    return this.editor.getLogger();
  }
  printLogs() {
    this.editor.printLogs();
  }
  async exportDocument() {
    return this.editor.export();
  }
  /** 导出为 Office 文件 Blob：Editor.bin → x2t → doc.{fileType}。 */
  async exportAsBlob() {
    if (this.editor.isOfficeXmlSizeLimitExceeded()) {
      const { file, fileName } = getDocumentObj(this.containerId);
      if (file) {
        return {
          blob: file,
          fileName: fileName || file.name,
          isOriginalFileFallback: true,
          fallbackReason: "officeXmlSizeLimitExceeded"
        };
      }
    }
    const binData = await this.editor.export();
    const exportFileType = binData.fileType || this.fileType.toLowerCase();
    const result = await convertBinToDocument(
      binData.binData,
      binData.fileName,
      exportFileType,
      binData.media,
      binData.themes,
      this.editor.getLogger()
    );
    return {
      blob: new Blob([result.data], {
        type: getOnlyOfficeMimeType(exportFileType)
      }),
      fileName: result.fileName
    };
  }
  /** 触发浏览器下载；内部先走 exportAsBlob 完整链路。 */
  async downloadExport() {
    const { blob, fileName } = await this.exportAsBlob();
    downloadBlob(blob, fileName);
  }
  /**
   * DSH patch: register the Ctrl/Cmd+S handler the editor should run.
   * Pass null to clear it (e.g. when the view is disposed or read-only).
   */
  setSaveShortcutHandler(handler) {
    this.editor.setSaveShortcutHandler(handler);
  }
  /**
   * DSH patch: register the auto-save trigger fired on document changes.
   * Pass null to clear it (read-only mode / disposal).
   */
  setChangeHandler(handler) {
    this.editor.setChangeHandler(handler);
  }
  /**
   * DSH patch: re-open the document inside the live editor (external change),
   * keeping the editor chrome instead of remounting the whole instance.
   * @returns whether the refresh was issued
   */
  async refreshDocument(file) {
    return this.editor.refreshDocument(file);
  }
  onLoadingChange(handler) {
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, handler);
    return () => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, handler);
    };
  }
  onOfficeXmlSizeLimitExceeded(handler) {
    onlyofficeEventbus.on(
      ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
      handler
    );
    return () => {
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
        handler
      );
    };
  }
  destroy() {
    this.editor.destroy();
    clearDocumentObj(this.containerId);
    this.ready = false;
  }
};
var OnlyOfficeManagerFactory = class {
  managers = /* @__PURE__ */ new Map();
  async open(options, document2) {
    const containerId = options.containerId ?? ONLYOFFICE_ID;
    let manager = this.managers.get(containerId);
    if (!manager) {
      await initializeOnlyOffice();
      const editor = editorManagerFactory.get(containerId);
      manager = OnlyOfficeManager.fromEditor(editor, {
        ...options,
        containerId
      });
      this.managers.set(containerId, manager);
    }
    await manager.openDocument({
      ...document2,
      readOnly: document2.readOnly ?? options.readOnly
    });
    return manager;
  }
  get(containerId) {
    return this.managers.get(containerId);
  }
  destroy(containerId) {
    this.managers.get(containerId)?.destroy();
    this.managers.delete(containerId);
  }
  destroyAll() {
    for (const manager of this.managers.values()) {
      manager.destroy();
    }
    this.managers.clear();
  }
};
var onlyOfficeManagerFactory = new OnlyOfficeManagerFactory();

// build/generated/x2t-worker-source.js
var X2T_WORKER_SOURCE = '(() => {\n  var __defProp = Object.defineProperty;\n  var __getOwnPropNames = Object.getOwnPropertyNames;\n  var __esm = (fn, res) => function __init() {\n    return fn && (res = (0, fn[__getOwnPropNames(fn)[0]])(fn = 0)), res;\n  };\n  var __export = (target, all) => {\n    for (var name in all)\n      __defProp(target, name, { get: all[name], enumerable: true });\n  };\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/streams.ts\n  function BrotliInput(buffer) {\n    this.buffer = buffer;\n    this.pos = 0;\n  }\n  function BrotliOutput(buf) {\n    this.buffer = buf;\n    this.pos = 0;\n  }\n  var init_streams = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/streams.ts"() {\n      BrotliInput.prototype.read = function(buf, i, count) {\n        if (this.pos + count > this.buffer.length) {\n          count = this.buffer.length - this.pos;\n        }\n        for (var p = 0; p < count; p++)\n          buf[i + p] = this.buffer[this.pos + p];\n        this.pos += count;\n        return count;\n      };\n      BrotliOutput.prototype.write = function(buf, count) {\n        if (this.pos + count > this.buffer.length)\n          throw new Error("Output buffer is not large enough");\n        this.buffer.set(buf.subarray(0, count), this.pos);\n        this.pos += count;\n        return count;\n      };\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/bit_reader.ts\n  function BrotliBitReader(input) {\n    this.buf_ = new Uint8Array(BROTLI_IBUF_SIZE);\n    this.input_ = input;\n    this.reset();\n  }\n  var BROTLI_READ_SIZE, BROTLI_IBUF_SIZE, BROTLI_IBUF_MASK, kBitMask, bit_reader_default;\n  var init_bit_reader = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/bit_reader.ts"() {\n      BROTLI_READ_SIZE = 4096;\n      BROTLI_IBUF_SIZE = 2 * BROTLI_READ_SIZE + 32;\n      BROTLI_IBUF_MASK = 2 * BROTLI_READ_SIZE - 1;\n      kBitMask = new Uint32Array([\n        0,\n        1,\n        3,\n        7,\n        15,\n        31,\n        63,\n        127,\n        255,\n        511,\n        1023,\n        2047,\n        4095,\n        8191,\n        16383,\n        32767,\n        65535,\n        131071,\n        262143,\n        524287,\n        1048575,\n        2097151,\n        4194303,\n        8388607,\n        16777215\n      ]);\n      BrotliBitReader.READ_SIZE = BROTLI_READ_SIZE;\n      BrotliBitReader.IBUF_MASK = BROTLI_IBUF_MASK;\n      BrotliBitReader.prototype.reset = function() {\n        this.buf_ptr_ = 0;\n        this.val_ = 0;\n        this.pos_ = 0;\n        this.bit_pos_ = 0;\n        this.bit_end_pos_ = 0;\n        this.eos_ = 0;\n        this.readMoreInput();\n        for (var i = 0; i < 4; i++) {\n          this.val_ |= this.buf_[this.pos_] << 8 * i;\n          ++this.pos_;\n        }\n        return this.bit_end_pos_ > 0;\n      };\n      BrotliBitReader.prototype.readMoreInput = function() {\n        if (this.bit_end_pos_ > 256) {\n          return;\n        } else if (this.eos_) {\n          if (this.bit_pos_ > this.bit_end_pos_)\n            throw new Error("Unexpected end of input " + this.bit_pos_ + " " + this.bit_end_pos_);\n        } else {\n          var dst = this.buf_ptr_;\n          var bytes_read = this.input_.read(this.buf_, dst, BROTLI_READ_SIZE);\n          if (bytes_read < 0) {\n            throw new Error("Unexpected end of input");\n          }\n          if (bytes_read < BROTLI_READ_SIZE) {\n            this.eos_ = 1;\n            for (var p = 0; p < 32; p++)\n              this.buf_[dst + bytes_read + p] = 0;\n          }\n          if (dst === 0) {\n            for (var p = 0; p < 32; p++)\n              this.buf_[(BROTLI_READ_SIZE << 1) + p] = this.buf_[p];\n            this.buf_ptr_ = BROTLI_READ_SIZE;\n          } else {\n            this.buf_ptr_ = 0;\n          }\n          this.bit_end_pos_ += bytes_read << 3;\n        }\n      };\n      BrotliBitReader.prototype.fillBitWindow = function() {\n        while (this.bit_pos_ >= 8) {\n          this.val_ >>>= 8;\n          this.val_ |= this.buf_[this.pos_ & BROTLI_IBUF_MASK] << 24;\n          ++this.pos_;\n          this.bit_pos_ = this.bit_pos_ - 8 >>> 0;\n          this.bit_end_pos_ = this.bit_end_pos_ - 8 >>> 0;\n        }\n      };\n      BrotliBitReader.prototype.readBits = function(n_bits) {\n        if (32 - this.bit_pos_ < n_bits) {\n          this.fillBitWindow();\n        }\n        var val = this.val_ >>> this.bit_pos_ & kBitMask[n_bits];\n        this.bit_pos_ += n_bits;\n        return val;\n      };\n      bit_reader_default = BrotliBitReader;\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/dictionary.bin.ts\n  var dictionary_bin_default;\n  var init_dictionary_bin = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/dictionary.bin.ts"() {\n      dictionary_bin_default = "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";\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/dictionary-browser.ts\n  function base64ToBytes(base64) {\n    const binary = atob(base64);\n    const bytes = new Uint8Array(binary.length);\n    for (let i = 0; i < binary.length; i++) {\n      bytes[i] = binary.charCodeAt(i);\n    }\n    return bytes;\n  }\n  function init() {\n    const compressed = base64ToBytes(dictionary_bin_default);\n    return BrotliDecompressBuffer(compressed);\n  }\n  var init_dictionary_browser = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/dictionary-browser.ts"() {\n      init_decode();\n      init_dictionary_bin();\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/dictionary.ts\n  function init2() {\n    dictionary = init();\n  }\n  var dictionary, offsetsByLength, sizeBitsByLength, minDictionaryWordLength, maxDictionaryWordLength;\n  var init_dictionary = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/dictionary.ts"() {\n      init_dictionary_browser();\n      offsetsByLength = new Uint32Array([\n        0,\n        0,\n        0,\n        0,\n        0,\n        4096,\n        9216,\n        21504,\n        35840,\n        44032,\n        53248,\n        63488,\n        74752,\n        87040,\n        93696,\n        100864,\n        104704,\n        106752,\n        108928,\n        113536,\n        115968,\n        118528,\n        119872,\n        121280,\n        122016\n      ]);\n      sizeBitsByLength = new Uint8Array([\n        0,\n        0,\n        0,\n        0,\n        10,\n        10,\n        11,\n        11,\n        10,\n        10,\n        10,\n        10,\n        10,\n        9,\n        9,\n        8,\n        7,\n        7,\n        8,\n        7,\n        7,\n        6,\n        6,\n        5,\n        5\n      ]);\n      minDictionaryWordLength = 4;\n      maxDictionaryWordLength = 24;\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/huffman.ts\n  function HuffmanCode(bits, value) {\n    this.bits = bits;\n    this.value = value;\n  }\n  function GetNextKey(key, len) {\n    var step = 1 << len - 1;\n    while (key & step) {\n      step >>= 1;\n    }\n    return (key & step - 1) + step;\n  }\n  function ReplicateValue(table, i, step, end, code) {\n    do {\n      end -= step;\n      table[i + end] = new HuffmanCode(code.bits, code.value);\n    } while (end > 0);\n  }\n  function NextTableBitSize(count, len, root_bits) {\n    var left = 1 << len - root_bits;\n    while (len < MAX_LENGTH) {\n      left -= count[len];\n      if (left <= 0) break;\n      ++len;\n      left <<= 1;\n    }\n    return len - root_bits;\n  }\n  function BrotliBuildHuffmanTable(root_table, table, root_bits, code_lengths, code_lengths_size) {\n    var start_table = table;\n    var code;\n    var len;\n    var symbol;\n    var key;\n    var step;\n    var low;\n    var mask;\n    var table_bits;\n    var table_size;\n    var total_size;\n    var sorted;\n    var count = new Int32Array(MAX_LENGTH + 1);\n    var offset = new Int32Array(MAX_LENGTH + 1);\n    sorted = new Int32Array(code_lengths_size);\n    for (symbol = 0; symbol < code_lengths_size; symbol++) {\n      count[code_lengths[symbol]]++;\n    }\n    offset[1] = 0;\n    for (len = 1; len < MAX_LENGTH; len++) {\n      offset[len + 1] = offset[len] + count[len];\n    }\n    for (symbol = 0; symbol < code_lengths_size; symbol++) {\n      if (code_lengths[symbol] !== 0) {\n        sorted[offset[code_lengths[symbol]]++] = symbol;\n      }\n    }\n    table_bits = root_bits;\n    table_size = 1 << table_bits;\n    total_size = table_size;\n    if (offset[MAX_LENGTH] === 1) {\n      for (key = 0; key < total_size; ++key) {\n        root_table[table + key] = new HuffmanCode(0, sorted[0] & 65535);\n      }\n      return total_size;\n    }\n    key = 0;\n    symbol = 0;\n    for (len = 1, step = 2; len <= root_bits; ++len, step <<= 1) {\n      for (; count[len] > 0; --count[len]) {\n        code = new HuffmanCode(len & 255, sorted[symbol++] & 65535);\n        ReplicateValue(root_table, table + key, step, table_size, code);\n        key = GetNextKey(key, len);\n      }\n    }\n    mask = total_size - 1;\n    low = -1;\n    for (len = root_bits + 1, step = 2; len <= MAX_LENGTH; ++len, step <<= 1) {\n      for (; count[len] > 0; --count[len]) {\n        if ((key & mask) !== low) {\n          table += table_size;\n          table_bits = NextTableBitSize(count, len, root_bits);\n          table_size = 1 << table_bits;\n          total_size += table_size;\n          low = key & mask;\n          root_table[start_table + low] = new HuffmanCode(table_bits + root_bits & 255, table - start_table - low & 65535);\n        }\n        code = new HuffmanCode(len - root_bits & 255, sorted[symbol++] & 65535);\n        ReplicateValue(root_table, table + (key >> root_bits), step, table_size, code);\n        key = GetNextKey(key, len);\n      }\n    }\n    return total_size;\n  }\n  var MAX_LENGTH;\n  var init_huffman = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/huffman.ts"() {\n      MAX_LENGTH = 15;\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/context.ts\n  var lookup, lookupOffsets;\n  var init_context = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/context.ts"() {\n      lookup = new Uint8Array([\n        /* CONTEXT_UTF8, last byte. */\n        /* ASCII range. */\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        4,\n        4,\n        0,\n        0,\n        4,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        8,\n        12,\n        16,\n        12,\n        12,\n        20,\n        12,\n        16,\n        24,\n        28,\n        12,\n        12,\n        32,\n        12,\n        36,\n        12,\n        44,\n        44,\n        44,\n        44,\n        44,\n        44,\n        44,\n        44,\n        44,\n        44,\n        32,\n        32,\n        24,\n        40,\n        28,\n        12,\n        12,\n        48,\n        52,\n        52,\n        52,\n        48,\n        52,\n        52,\n        52,\n        48,\n        52,\n        52,\n        52,\n        52,\n        52,\n        48,\n        52,\n        52,\n        52,\n        52,\n        52,\n        48,\n        52,\n        52,\n        52,\n        52,\n        52,\n        24,\n        12,\n        28,\n        12,\n        12,\n        12,\n        56,\n        60,\n        60,\n        60,\n        56,\n        60,\n        60,\n        60,\n        56,\n        60,\n        60,\n        60,\n        60,\n        60,\n        56,\n        60,\n        60,\n        60,\n        60,\n        60,\n        56,\n        60,\n        60,\n        60,\n        60,\n        60,\n        24,\n        12,\n        28,\n        12,\n        0,\n        /* UTF8 continuation byte range. */\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        0,\n        1,\n        /* UTF8 lead byte range. */\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        2,\n        3,\n        /* CONTEXT_UTF8 second last byte. */\n        /* ASCII range. */\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        1,\n        1,\n        1,\n        1,\n        0,\n        /* UTF8 continuation byte range. */\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        /* UTF8 lead byte range. */\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        /* CONTEXT_SIGNED, second last byte. */\n        0,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        1,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        4,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        5,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        6,\n        7,\n        /* CONTEXT_SIGNED, last byte, same as the above values shifted by 3 bits. */\n        0,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        8,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        16,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        24,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        32,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        40,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        48,\n        56,\n        /* CONTEXT_LSB6, last byte. */\n        0,\n        1,\n        2,\n        3,\n        4,\n        5,\n        6,\n        7,\n        8,\n        9,\n        10,\n        11,\n        12,\n        13,\n        14,\n        15,\n        16,\n        17,\n        18,\n        19,\n        20,\n        21,\n        22,\n        23,\n        24,\n        25,\n        26,\n        27,\n        28,\n        29,\n        30,\n        31,\n        32,\n        33,\n        34,\n        35,\n        36,\n        37,\n        38,\n        39,\n        40,\n        41,\n        42,\n        43,\n        44,\n        45,\n        46,\n        47,\n        48,\n        49,\n        50,\n        51,\n        52,\n        53,\n        54,\n        55,\n        56,\n        57,\n        58,\n        59,\n        60,\n        61,\n        62,\n        63,\n        0,\n        1,\n        2,\n        3,\n        4,\n        5,\n        6,\n        7,\n        8,\n        9,\n        10,\n        11,\n        12,\n        13,\n        14,\n        15,\n        16,\n        17,\n        18,\n        19,\n        20,\n        21,\n        22,\n        23,\n        24,\n        25,\n        26,\n        27,\n        28,\n        29,\n        30,\n        31,\n        32,\n        33,\n        34,\n        35,\n        36,\n        37,\n        38,\n        39,\n        40,\n        41,\n        42,\n        43,\n        44,\n        45,\n        46,\n        47,\n        48,\n        49,\n        50,\n        51,\n        52,\n        53,\n        54,\n        55,\n        56,\n        57,\n        58,\n        59,\n        60,\n        61,\n        62,\n        63,\n        0,\n        1,\n        2,\n        3,\n        4,\n        5,\n        6,\n        7,\n        8,\n        9,\n        10,\n        11,\n        12,\n        13,\n        14,\n        15,\n        16,\n        17,\n        18,\n        19,\n        20,\n        21,\n        22,\n        23,\n        24,\n        25,\n        26,\n        27,\n        28,\n        29,\n        30,\n        31,\n        32,\n        33,\n        34,\n        35,\n        36,\n        37,\n        38,\n        39,\n        40,\n        41,\n        42,\n        43,\n        44,\n        45,\n        46,\n        47,\n        48,\n        49,\n        50,\n        51,\n        52,\n        53,\n        54,\n        55,\n        56,\n        57,\n        58,\n        59,\n        60,\n        61,\n        62,\n        63,\n        0,\n        1,\n        2,\n        3,\n        4,\n        5,\n        6,\n        7,\n        8,\n        9,\n        10,\n        11,\n        12,\n        13,\n        14,\n        15,\n        16,\n        17,\n        18,\n        19,\n        20,\n        21,\n        22,\n        23,\n        24,\n        25,\n        26,\n        27,\n        28,\n        29,\n        30,\n        31,\n        32,\n        33,\n        34,\n        35,\n        36,\n        37,\n        38,\n        39,\n        40,\n        41,\n        42,\n        43,\n        44,\n        45,\n        46,\n        47,\n        48,\n        49,\n        50,\n        51,\n        52,\n        53,\n        54,\n        55,\n        56,\n        57,\n        58,\n        59,\n        60,\n        61,\n        62,\n        63,\n        /* CONTEXT_MSB6, last byte. */\n        0,\n        0,\n        0,\n        0,\n        1,\n        1,\n        1,\n        1,\n        2,\n        2,\n        2,\n        2,\n        3,\n        3,\n        3,\n        3,\n        4,\n        4,\n        4,\n        4,\n        5,\n        5,\n        5,\n        5,\n        6,\n        6,\n        6,\n        6,\n        7,\n        7,\n        7,\n        7,\n        8,\n        8,\n        8,\n        8,\n        9,\n        9,\n        9,\n        9,\n        10,\n        10,\n        10,\n        10,\n        11,\n        11,\n        11,\n        11,\n        12,\n        12,\n        12,\n        12,\n        13,\n        13,\n        13,\n        13,\n        14,\n        14,\n        14,\n        14,\n        15,\n        15,\n        15,\n        15,\n        16,\n        16,\n        16,\n        16,\n        17,\n        17,\n        17,\n        17,\n        18,\n        18,\n        18,\n        18,\n        19,\n        19,\n        19,\n        19,\n        20,\n        20,\n        20,\n        20,\n        21,\n        21,\n        21,\n        21,\n        22,\n        22,\n        22,\n        22,\n        23,\n        23,\n        23,\n        23,\n        24,\n        24,\n        24,\n        24,\n        25,\n        25,\n        25,\n        25,\n        26,\n        26,\n        26,\n        26,\n        27,\n        27,\n        27,\n        27,\n        28,\n        28,\n        28,\n        28,\n        29,\n        29,\n        29,\n        29,\n        30,\n        30,\n        30,\n        30,\n        31,\n        31,\n        31,\n        31,\n        32,\n        32,\n        32,\n        32,\n        33,\n        33,\n        33,\n        33,\n        34,\n        34,\n        34,\n        34,\n        35,\n        35,\n        35,\n        35,\n        36,\n        36,\n        36,\n        36,\n        37,\n        37,\n        37,\n        37,\n        38,\n        38,\n        38,\n        38,\n        39,\n        39,\n        39,\n        39,\n        40,\n        40,\n        40,\n        40,\n        41,\n        41,\n        41,\n        41,\n        42,\n        42,\n        42,\n        42,\n        43,\n        43,\n        43,\n        43,\n        44,\n        44,\n        44,\n        44,\n        45,\n        45,\n        45,\n        45,\n        46,\n        46,\n        46,\n        46,\n        47,\n        47,\n        47,\n        47,\n        48,\n        48,\n        48,\n        48,\n        49,\n        49,\n        49,\n        49,\n        50,\n        50,\n        50,\n        50,\n        51,\n        51,\n        51,\n        51,\n        52,\n        52,\n        52,\n        52,\n        53,\n        53,\n        53,\n        53,\n        54,\n        54,\n        54,\n        54,\n        55,\n        55,\n        55,\n        55,\n        56,\n        56,\n        56,\n        56,\n        57,\n        57,\n        57,\n        57,\n        58,\n        58,\n        58,\n        58,\n        59,\n        59,\n        59,\n        59,\n        60,\n        60,\n        60,\n        60,\n        61,\n        61,\n        61,\n        61,\n        62,\n        62,\n        62,\n        62,\n        63,\n        63,\n        63,\n        63,\n        /* CONTEXT_{M,L}SB6, second last byte, */\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0,\n        0\n      ]);\n      lookupOffsets = new Uint16Array([\n        /* CONTEXT_LSB6 */\n        1024,\n        1536,\n        /* CONTEXT_MSB6 */\n        1280,\n        1536,\n        /* CONTEXT_UTF8 */\n        0,\n        256,\n        /* CONTEXT_SIGNED */\n        768,\n        512\n      ]);\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/prefix.ts\n  function PrefixCodeRange(offset, nbits) {\n    this.offset = offset;\n    this.nbits = nbits;\n  }\n  var kBlockLengthPrefixCode, kInsertLengthPrefixCode, kCopyLengthPrefixCode, kInsertRangeLut, kCopyRangeLut;\n  var init_prefix = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/prefix.ts"() {\n      kBlockLengthPrefixCode = [\n        new PrefixCodeRange(1, 2),\n        new PrefixCodeRange(5, 2),\n        new PrefixCodeRange(9, 2),\n        new PrefixCodeRange(13, 2),\n        new PrefixCodeRange(17, 3),\n        new PrefixCodeRange(25, 3),\n        new PrefixCodeRange(33, 3),\n        new PrefixCodeRange(41, 3),\n        new PrefixCodeRange(49, 4),\n        new PrefixCodeRange(65, 4),\n        new PrefixCodeRange(81, 4),\n        new PrefixCodeRange(97, 4),\n        new PrefixCodeRange(113, 5),\n        new PrefixCodeRange(145, 5),\n        new PrefixCodeRange(177, 5),\n        new PrefixCodeRange(209, 5),\n        new PrefixCodeRange(241, 6),\n        new PrefixCodeRange(305, 6),\n        new PrefixCodeRange(369, 7),\n        new PrefixCodeRange(497, 8),\n        new PrefixCodeRange(753, 9),\n        new PrefixCodeRange(1265, 10),\n        new PrefixCodeRange(2289, 11),\n        new PrefixCodeRange(4337, 12),\n        new PrefixCodeRange(8433, 13),\n        new PrefixCodeRange(16625, 24)\n      ];\n      kInsertLengthPrefixCode = [\n        new PrefixCodeRange(0, 0),\n        new PrefixCodeRange(1, 0),\n        new PrefixCodeRange(2, 0),\n        new PrefixCodeRange(3, 0),\n        new PrefixCodeRange(4, 0),\n        new PrefixCodeRange(5, 0),\n        new PrefixCodeRange(6, 1),\n        new PrefixCodeRange(8, 1),\n        new PrefixCodeRange(10, 2),\n        new PrefixCodeRange(14, 2),\n        new PrefixCodeRange(18, 3),\n        new PrefixCodeRange(26, 3),\n        new PrefixCodeRange(34, 4),\n        new PrefixCodeRange(50, 4),\n        new PrefixCodeRange(66, 5),\n        new PrefixCodeRange(98, 5),\n        new PrefixCodeRange(130, 6),\n        new PrefixCodeRange(194, 7),\n        new PrefixCodeRange(322, 8),\n        new PrefixCodeRange(578, 9),\n        new PrefixCodeRange(1090, 10),\n        new PrefixCodeRange(2114, 12),\n        new PrefixCodeRange(6210, 14),\n        new PrefixCodeRange(22594, 24)\n      ];\n      kCopyLengthPrefixCode = [\n        new PrefixCodeRange(2, 0),\n        new PrefixCodeRange(3, 0),\n        new PrefixCodeRange(4, 0),\n        new PrefixCodeRange(5, 0),\n        new PrefixCodeRange(6, 0),\n        new PrefixCodeRange(7, 0),\n        new PrefixCodeRange(8, 0),\n        new PrefixCodeRange(9, 0),\n        new PrefixCodeRange(10, 1),\n        new PrefixCodeRange(12, 1),\n        new PrefixCodeRange(14, 2),\n        new PrefixCodeRange(18, 2),\n        new PrefixCodeRange(22, 3),\n        new PrefixCodeRange(30, 3),\n        new PrefixCodeRange(38, 4),\n        new PrefixCodeRange(54, 4),\n        new PrefixCodeRange(70, 5),\n        new PrefixCodeRange(102, 5),\n        new PrefixCodeRange(134, 6),\n        new PrefixCodeRange(198, 7),\n        new PrefixCodeRange(326, 8),\n        new PrefixCodeRange(582, 9),\n        new PrefixCodeRange(1094, 10),\n        new PrefixCodeRange(2118, 24)\n      ];\n      kInsertRangeLut = [\n        0,\n        0,\n        8,\n        8,\n        0,\n        16,\n        8,\n        16,\n        16\n      ];\n      kCopyRangeLut = [\n        0,\n        8,\n        0,\n        8,\n        16,\n        0,\n        16,\n        8,\n        16\n      ];\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/transform.ts\n  function Transform(prefix, transform, suffix) {\n    this.prefix = new Uint8Array(prefix.length);\n    this.transform = transform;\n    this.suffix = new Uint8Array(suffix.length);\n    for (var i = 0; i < prefix.length; i++)\n      this.prefix[i] = prefix.charCodeAt(i);\n    for (var i = 0; i < suffix.length; i++)\n      this.suffix[i] = suffix.charCodeAt(i);\n  }\n  function ToUpperCase(p, i) {\n    if (p[i] < 192) {\n      if (p[i] >= 97 && p[i] <= 122) {\n        p[i] ^= 32;\n      }\n      return 1;\n    }\n    if (p[i] < 224) {\n      p[i + 1] ^= 32;\n      return 2;\n    }\n    p[i + 2] ^= 5;\n    return 3;\n  }\n  function transformDictionaryWord(dst, idx, word, len, transform) {\n    var prefix = kTransforms[transform].prefix;\n    var suffix = kTransforms[transform].suffix;\n    var t = kTransforms[transform].transform;\n    var skip = t < kOmitFirst1 ? 0 : t - (kOmitFirst1 - 1);\n    var i = 0;\n    var start_idx = idx;\n    var uppercase;\n    if (skip > len) {\n      skip = len;\n    }\n    var prefix_pos = 0;\n    while (prefix_pos < prefix.length) {\n      dst[idx++] = prefix[prefix_pos++];\n    }\n    word += skip;\n    len -= skip;\n    if (t <= kOmitLast9) {\n      len -= t;\n    }\n    for (i = 0; i < len; i++) {\n      dst[idx++] = dictionary[word + i];\n    }\n    uppercase = idx - len;\n    if (t === kUppercaseFirst) {\n      ToUpperCase(dst, uppercase);\n    } else if (t === kUppercaseAll) {\n      while (len > 0) {\n        var step = ToUpperCase(dst, uppercase);\n        uppercase += step;\n        len -= step;\n      }\n    }\n    var suffix_pos = 0;\n    while (suffix_pos < suffix.length) {\n      dst[idx++] = suffix[suffix_pos++];\n    }\n    return idx - start_idx;\n  }\n  var kIdentity, kOmitLast1, kOmitLast2, kOmitLast3, kOmitLast4, kOmitLast5, kOmitLast6, kOmitLast7, kOmitLast8, kOmitLast9, kUppercaseFirst, kUppercaseAll, kOmitFirst1, kOmitFirst2, kOmitFirst3, kOmitFirst4, kOmitFirst5, kOmitFirst6, kOmitFirst7, kOmitFirst9, kTransforms, kNumTransforms;\n  var init_transform = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/transform.ts"() {\n      init_dictionary();\n      kIdentity = 0;\n      kOmitLast1 = 1;\n      kOmitLast2 = 2;\n      kOmitLast3 = 3;\n      kOmitLast4 = 4;\n      kOmitLast5 = 5;\n      kOmitLast6 = 6;\n      kOmitLast7 = 7;\n      kOmitLast8 = 8;\n      kOmitLast9 = 9;\n      kUppercaseFirst = 10;\n      kUppercaseAll = 11;\n      kOmitFirst1 = 12;\n      kOmitFirst2 = 13;\n      kOmitFirst3 = 14;\n      kOmitFirst4 = 15;\n      kOmitFirst5 = 16;\n      kOmitFirst6 = 17;\n      kOmitFirst7 = 18;\n      kOmitFirst9 = 20;\n      kTransforms = [\n        new Transform("", kIdentity, ""),\n        new Transform("", kIdentity, " "),\n        new Transform(" ", kIdentity, " "),\n        new Transform("", kOmitFirst1, ""),\n        new Transform("", kUppercaseFirst, " "),\n        new Transform("", kIdentity, " the "),\n        new Transform(" ", kIdentity, ""),\n        new Transform("s ", kIdentity, " "),\n        new Transform("", kIdentity, " of "),\n        new Transform("", kUppercaseFirst, ""),\n        new Transform("", kIdentity, " and "),\n        new Transform("", kOmitFirst2, ""),\n        new Transform("", kOmitLast1, ""),\n        new Transform(", ", kIdentity, " "),\n        new Transform("", kIdentity, ", "),\n        new Transform(" ", kUppercaseFirst, " "),\n        new Transform("", kIdentity, " in "),\n        new Transform("", kIdentity, " to "),\n        new Transform("e ", kIdentity, " "),\n        new Transform("", kIdentity, \'"\'),\n        new Transform("", kIdentity, "."),\n        new Transform("", kIdentity, \'">\'),\n        new Transform("", kIdentity, "\\n"),\n        new Transform("", kOmitLast3, ""),\n        new Transform("", kIdentity, "]"),\n        new Transform("", kIdentity, " for "),\n        new Transform("", kOmitFirst3, ""),\n        new Transform("", kOmitLast2, ""),\n        new Transform("", kIdentity, " a "),\n        new Transform("", kIdentity, " that "),\n        new Transform(" ", kUppercaseFirst, ""),\n        new Transform("", kIdentity, ". "),\n        new Transform(".", kIdentity, ""),\n        new Transform(" ", kIdentity, ", "),\n        new Transform("", kOmitFirst4, ""),\n        new Transform("", kIdentity, " with "),\n        new Transform("", kIdentity, "\'"),\n        new Transform("", kIdentity, " from "),\n        new Transform("", kIdentity, " by "),\n        new Transform("", kOmitFirst5, ""),\n        new Transform("", kOmitFirst6, ""),\n        new Transform(" the ", kIdentity, ""),\n        new Transform("", kOmitLast4, ""),\n        new Transform("", kIdentity, ". The "),\n        new Transform("", kUppercaseAll, ""),\n        new Transform("", kIdentity, " on "),\n        new Transform("", kIdentity, " as "),\n        new Transform("", kIdentity, " is "),\n        new Transform("", kOmitLast7, ""),\n        new Transform("", kOmitLast1, "ing "),\n        new Transform("", kIdentity, "\\n	"),\n        new Transform("", kIdentity, ":"),\n        new Transform(" ", kIdentity, ". "),\n        new Transform("", kIdentity, "ed "),\n        new Transform("", kOmitFirst9, ""),\n        new Transform("", kOmitFirst7, ""),\n        new Transform("", kOmitLast6, ""),\n        new Transform("", kIdentity, "("),\n        new Transform("", kUppercaseFirst, ", "),\n        new Transform("", kOmitLast8, ""),\n        new Transform("", kIdentity, " at "),\n        new Transform("", kIdentity, "ly "),\n        new Transform(" the ", kIdentity, " of "),\n        new Transform("", kOmitLast5, ""),\n        new Transform("", kOmitLast9, ""),\n        new Transform(" ", kUppercaseFirst, ", "),\n        new Transform("", kUppercaseFirst, \'"\'),\n        new Transform(".", kIdentity, "("),\n        new Transform("", kUppercaseAll, " "),\n        new Transform("", kUppercaseFirst, \'">\'),\n        new Transform("", kIdentity, \'="\'),\n        new Transform(" ", kIdentity, "."),\n        new Transform(".com/", kIdentity, ""),\n        new Transform(" the ", kIdentity, " of the "),\n        new Transform("", kUppercaseFirst, "\'"),\n        new Transform("", kIdentity, ". This "),\n        new Transform("", kIdentity, ","),\n        new Transform(".", kIdentity, " "),\n        new Transform("", kUppercaseFirst, "("),\n        new Transform("", kUppercaseFirst, "."),\n        new Transform("", kIdentity, " not "),\n        new Transform(" ", kIdentity, \'="\'),\n        new Transform("", kIdentity, "er "),\n        new Transform(" ", kUppercaseAll, " "),\n        new Transform("", kIdentity, "al "),\n        new Transform(" ", kUppercaseAll, ""),\n        new Transform("", kIdentity, "=\'"),\n        new Transform("", kUppercaseAll, \'"\'),\n        new Transform("", kUppercaseFirst, ". "),\n        new Transform(" ", kIdentity, "("),\n        new Transform("", kIdentity, "ful "),\n        new Transform(" ", kUppercaseFirst, ". "),\n        new Transform("", kIdentity, "ive "),\n        new Transform("", kIdentity, "less "),\n        new Transform("", kUppercaseAll, "\'"),\n        new Transform("", kIdentity, "est "),\n        new Transform(" ", kUppercaseFirst, "."),\n        new Transform("", kUppercaseAll, \'">\'),\n        new Transform(" ", kIdentity, "=\'"),\n        new Transform("", kUppercaseFirst, ","),\n        new Transform("", kIdentity, "ize "),\n        new Transform("", kUppercaseAll, "."),\n        new Transform("Â ", kIdentity, ""),\n        new Transform(" ", kIdentity, ","),\n        new Transform("", kUppercaseFirst, \'="\'),\n        new Transform("", kUppercaseAll, \'="\'),\n        new Transform("", kIdentity, "ous "),\n        new Transform("", kUppercaseAll, ", "),\n        new Transform("", kUppercaseFirst, "=\'"),\n        new Transform(" ", kUppercaseFirst, ","),\n        new Transform(" ", kUppercaseAll, \'="\'),\n        new Transform(" ", kUppercaseAll, ", "),\n        new Transform("", kUppercaseAll, ","),\n        new Transform("", kUppercaseAll, "("),\n        new Transform("", kUppercaseAll, ". "),\n        new Transform(" ", kUppercaseAll, "."),\n        new Transform("", kUppercaseAll, "=\'"),\n        new Transform(" ", kUppercaseAll, ". "),\n        new Transform(" ", kUppercaseFirst, \'="\'),\n        new Transform(" ", kUppercaseAll, "=\'"),\n        new Transform(" ", kUppercaseFirst, "=\'")\n      ];\n      kNumTransforms = kTransforms.length;\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/decode.ts\n  function DecodeWindowBits(br) {\n    var n;\n    if (br.readBits(1) === 0) {\n      return 16;\n    }\n    n = br.readBits(3);\n    if (n > 0) {\n      return 17 + n;\n    }\n    n = br.readBits(3);\n    if (n > 0) {\n      return 8 + n;\n    }\n    return 17;\n  }\n  function DecodeVarLenUint8(br) {\n    if (br.readBits(1)) {\n      var nbits = br.readBits(3);\n      if (nbits === 0) {\n        return 1;\n      } else {\n        return br.readBits(nbits) + (1 << nbits);\n      }\n    }\n    return 0;\n  }\n  function MetaBlockLength() {\n    this.meta_block_length = 0;\n    this.input_end = 0;\n    this.is_uncompressed = 0;\n    this.is_metadata = false;\n  }\n  function DecodeMetaBlockLength(br) {\n    var out = new MetaBlockLength();\n    var size_nibbles;\n    var size_bytes;\n    var i;\n    out.input_end = br.readBits(1);\n    if (out.input_end && br.readBits(1)) {\n      return out;\n    }\n    size_nibbles = br.readBits(2) + 4;\n    if (size_nibbles === 7) {\n      out.is_metadata = true;\n      if (br.readBits(1) !== 0)\n        throw new Error("Invalid reserved bit");\n      size_bytes = br.readBits(2);\n      if (size_bytes === 0)\n        return out;\n      for (i = 0; i < size_bytes; i++) {\n        var next_byte = br.readBits(8);\n        if (i + 1 === size_bytes && size_bytes > 1 && next_byte === 0)\n          throw new Error("Invalid size byte");\n        out.meta_block_length |= next_byte << i * 8;\n      }\n    } else {\n      for (i = 0; i < size_nibbles; ++i) {\n        var next_nibble = br.readBits(4);\n        if (i + 1 === size_nibbles && size_nibbles > 4 && next_nibble === 0)\n          throw new Error("Invalid size nibble");\n        out.meta_block_length |= next_nibble << i * 4;\n      }\n    }\n    ++out.meta_block_length;\n    if (!out.input_end && !out.is_metadata) {\n      out.is_uncompressed = br.readBits(1);\n    }\n    return out;\n  }\n  function ReadSymbol(table, index, br) {\n    var start_index = index;\n    var nbits;\n    br.fillBitWindow();\n    index += br.val_ >>> br.bit_pos_ & HUFFMAN_TABLE_MASK;\n    nbits = table[index].bits - HUFFMAN_TABLE_BITS;\n    if (nbits > 0) {\n      br.bit_pos_ += HUFFMAN_TABLE_BITS;\n      index += table[index].value;\n      index += br.val_ >>> br.bit_pos_ & (1 << nbits) - 1;\n    }\n    br.bit_pos_ += table[index].bits;\n    return table[index].value;\n  }\n  function ReadHuffmanCodeLengths(code_length_code_lengths, num_symbols, code_lengths, br) {\n    var symbol = 0;\n    var prev_code_len = kDefaultCodeLength;\n    var repeat = 0;\n    var repeat_code_len = 0;\n    var space = 32768;\n    var table = [];\n    for (var i = 0; i < 32; i++)\n      table.push(new HuffmanCode(0, 0));\n    BrotliBuildHuffmanTable(table, 0, 5, code_length_code_lengths, CODE_LENGTH_CODES);\n    while (symbol < num_symbols && space > 0) {\n      var p = 0;\n      var code_len;\n      br.readMoreInput();\n      br.fillBitWindow();\n      p += br.val_ >>> br.bit_pos_ & 31;\n      br.bit_pos_ += table[p].bits;\n      code_len = table[p].value & 255;\n      if (code_len < kCodeLengthRepeatCode) {\n        repeat = 0;\n        code_lengths[symbol++] = code_len;\n        if (code_len !== 0) {\n          prev_code_len = code_len;\n          space -= 32768 >> code_len;\n        }\n      } else {\n        var extra_bits = code_len - 14;\n        var old_repeat;\n        var repeat_delta;\n        var new_len = 0;\n        if (code_len === kCodeLengthRepeatCode) {\n          new_len = prev_code_len;\n        }\n        if (repeat_code_len !== new_len) {\n          repeat = 0;\n          repeat_code_len = new_len;\n        }\n        old_repeat = repeat;\n        if (repeat > 0) {\n          repeat -= 2;\n          repeat <<= extra_bits;\n        }\n        repeat += br.readBits(extra_bits) + 3;\n        repeat_delta = repeat - old_repeat;\n        if (symbol + repeat_delta > num_symbols) {\n          throw new Error("[ReadHuffmanCodeLengths] symbol + repeat_delta > num_symbols");\n        }\n        for (var x = 0; x < repeat_delta; x++)\n          code_lengths[symbol + x] = repeat_code_len;\n        symbol += repeat_delta;\n        if (repeat_code_len !== 0) {\n          space -= repeat_delta << 15 - repeat_code_len;\n        }\n      }\n    }\n    if (space !== 0) {\n      throw new Error("[ReadHuffmanCodeLengths] space = " + space);\n    }\n    for (; symbol < num_symbols; symbol++)\n      code_lengths[symbol] = 0;\n  }\n  function ReadHuffmanCode(alphabet_size, tables, table, br) {\n    var table_size = 0;\n    var simple_code_or_skip;\n    var code_lengths = new Uint8Array(alphabet_size);\n    br.readMoreInput();\n    simple_code_or_skip = br.readBits(2);\n    if (simple_code_or_skip === 1) {\n      var i;\n      var max_bits_counter = alphabet_size - 1;\n      var max_bits = 0;\n      var symbols = new Int32Array(4);\n      var num_symbols = br.readBits(2) + 1;\n      while (max_bits_counter) {\n        max_bits_counter >>= 1;\n        ++max_bits;\n      }\n      for (i = 0; i < num_symbols; ++i) {\n        symbols[i] = br.readBits(max_bits) % alphabet_size;\n        code_lengths[symbols[i]] = 2;\n      }\n      code_lengths[symbols[0]] = 1;\n      switch (num_symbols) {\n        case 1:\n          break;\n        case 3:\n          if (symbols[0] === symbols[1] || symbols[0] === symbols[2] || symbols[1] === symbols[2]) {\n            throw new Error("[ReadHuffmanCode] invalid symbols");\n          }\n          break;\n        case 2:\n          if (symbols[0] === symbols[1]) {\n            throw new Error("[ReadHuffmanCode] invalid symbols");\n          }\n          code_lengths[symbols[1]] = 1;\n          break;\n        case 4:\n          if (symbols[0] === symbols[1] || symbols[0] === symbols[2] || symbols[0] === symbols[3] || symbols[1] === symbols[2] || symbols[1] === symbols[3] || symbols[2] === symbols[3]) {\n            throw new Error("[ReadHuffmanCode] invalid symbols");\n          }\n          if (br.readBits(1)) {\n            code_lengths[symbols[2]] = 3;\n            code_lengths[symbols[3]] = 3;\n          } else {\n            code_lengths[symbols[0]] = 2;\n          }\n          break;\n      }\n    } else {\n      var i;\n      var code_length_code_lengths = new Uint8Array(CODE_LENGTH_CODES);\n      var space = 32;\n      var num_codes = 0;\n      var huff = [\n        new HuffmanCode(2, 0),\n        new HuffmanCode(2, 4),\n        new HuffmanCode(2, 3),\n        new HuffmanCode(3, 2),\n        new HuffmanCode(2, 0),\n        new HuffmanCode(2, 4),\n        new HuffmanCode(2, 3),\n        new HuffmanCode(4, 1),\n        new HuffmanCode(2, 0),\n        new HuffmanCode(2, 4),\n        new HuffmanCode(2, 3),\n        new HuffmanCode(3, 2),\n        new HuffmanCode(2, 0),\n        new HuffmanCode(2, 4),\n        new HuffmanCode(2, 3),\n        new HuffmanCode(4, 5)\n      ];\n      for (i = simple_code_or_skip; i < CODE_LENGTH_CODES && space > 0; ++i) {\n        var code_len_idx = kCodeLengthCodeOrder[i];\n        var p = 0;\n        var v;\n        br.fillBitWindow();\n        p += br.val_ >>> br.bit_pos_ & 15;\n        br.bit_pos_ += huff[p].bits;\n        v = huff[p].value;\n        code_length_code_lengths[code_len_idx] = v;\n        if (v !== 0) {\n          space -= 32 >> v;\n          ++num_codes;\n        }\n      }\n      if (!(num_codes === 1 || space === 0))\n        throw new Error("[ReadHuffmanCode] invalid num_codes or space");\n      ReadHuffmanCodeLengths(code_length_code_lengths, alphabet_size, code_lengths, br);\n    }\n    table_size = BrotliBuildHuffmanTable(tables, table, HUFFMAN_TABLE_BITS, code_lengths, alphabet_size);\n    if (table_size === 0) {\n      throw new Error("[ReadHuffmanCode] BuildHuffmanTable failed: ");\n    }\n    return table_size;\n  }\n  function ReadBlockLength(table, index, br) {\n    var code;\n    var nbits;\n    code = ReadSymbol(table, index, br);\n    nbits = kBlockLengthPrefixCode[code].nbits;\n    return kBlockLengthPrefixCode[code].offset + br.readBits(nbits);\n  }\n  function TranslateShortCodes(code, ringbuffer, index) {\n    var val;\n    if (code < NUM_DISTANCE_SHORT_CODES) {\n      index += kDistanceShortCodeIndexOffset[code];\n      index &= 3;\n      val = ringbuffer[index] + kDistanceShortCodeValueOffset[code];\n    } else {\n      val = code - NUM_DISTANCE_SHORT_CODES + 1;\n    }\n    return val;\n  }\n  function MoveToFront(v, index) {\n    var value = v[index];\n    var i = index;\n    for (; i; --i) v[i] = v[i - 1];\n    v[0] = value;\n  }\n  function InverseMoveToFrontTransform(v, v_len) {\n    var mtf = new Uint8Array(256);\n    var i;\n    for (i = 0; i < 256; ++i) {\n      mtf[i] = i;\n    }\n    for (i = 0; i < v_len; ++i) {\n      var index = v[i];\n      v[i] = mtf[index];\n      if (index) MoveToFront(mtf, index);\n    }\n  }\n  function HuffmanTreeGroup(alphabet_size, num_htrees) {\n    this.alphabet_size = alphabet_size;\n    this.num_htrees = num_htrees;\n    this.codes = new Array(num_htrees + num_htrees * kMaxHuffmanTableSize[alphabet_size + 31 >>> 5]);\n    this.htrees = new Uint32Array(num_htrees);\n  }\n  function DecodeContextMap(context_map_size, br) {\n    var out = { num_htrees: null, context_map: null };\n    var use_rle_for_zeros;\n    var max_run_length_prefix = 0;\n    var table;\n    var i;\n    br.readMoreInput();\n    var num_htrees = out.num_htrees = DecodeVarLenUint8(br) + 1;\n    var context_map = out.context_map = new Uint8Array(context_map_size);\n    if (num_htrees <= 1) {\n      return out;\n    }\n    use_rle_for_zeros = br.readBits(1);\n    if (use_rle_for_zeros) {\n      max_run_length_prefix = br.readBits(4) + 1;\n    }\n    table = [];\n    for (i = 0; i < HUFFMAN_MAX_TABLE_SIZE; i++) {\n      table[i] = new HuffmanCode(0, 0);\n    }\n    ReadHuffmanCode(num_htrees + max_run_length_prefix, table, 0, br);\n    for (i = 0; i < context_map_size; ) {\n      var code;\n      br.readMoreInput();\n      code = ReadSymbol(table, 0, br);\n      if (code === 0) {\n        context_map[i] = 0;\n        ++i;\n      } else if (code <= max_run_length_prefix) {\n        var reps = 1 + (1 << code) + br.readBits(code);\n        while (--reps) {\n          if (i >= context_map_size) {\n            throw new Error("[DecodeContextMap] i >= context_map_size");\n          }\n          context_map[i] = 0;\n          ++i;\n        }\n      } else {\n        context_map[i] = code - max_run_length_prefix;\n        ++i;\n      }\n    }\n    if (br.readBits(1)) {\n      InverseMoveToFrontTransform(context_map, context_map_size);\n    }\n    return out;\n  }\n  function DecodeBlockType(max_block_type, trees, tree_type, block_types, ringbuffers, indexes, br) {\n    var ringbuffer = tree_type * 2;\n    var index = tree_type;\n    var type_code = ReadSymbol(trees, tree_type * HUFFMAN_MAX_TABLE_SIZE, br);\n    var block_type;\n    if (type_code === 0) {\n      block_type = ringbuffers[ringbuffer + (indexes[index] & 1)];\n    } else if (type_code === 1) {\n      block_type = ringbuffers[ringbuffer + (indexes[index] - 1 & 1)] + 1;\n    } else {\n      block_type = type_code - 2;\n    }\n    if (block_type >= max_block_type) {\n      block_type -= max_block_type;\n    }\n    block_types[tree_type] = block_type;\n    ringbuffers[ringbuffer + (indexes[index] & 1)] = block_type;\n    ++indexes[index];\n  }\n  function CopyUncompressedBlockToOutput(output, len, pos, ringbuffer, ringbuffer_mask, br) {\n    var rb_size = ringbuffer_mask + 1;\n    var rb_pos = pos & ringbuffer_mask;\n    var br_pos = br.pos_ & bit_reader_default.IBUF_MASK;\n    var nbytes;\n    if (len < 8 || br.bit_pos_ + (len << 3) < br.bit_end_pos_) {\n      while (len-- > 0) {\n        br.readMoreInput();\n        ringbuffer[rb_pos++] = br.readBits(8);\n        if (rb_pos === rb_size) {\n          output.write(ringbuffer, rb_size);\n          rb_pos = 0;\n        }\n      }\n      return;\n    }\n    if (br.bit_end_pos_ < 32) {\n      throw new Error("[CopyUncompressedBlockToOutput] br.bit_end_pos_ < 32");\n    }\n    while (br.bit_pos_ < 32) {\n      ringbuffer[rb_pos] = br.val_ >>> br.bit_pos_;\n      br.bit_pos_ += 8;\n      ++rb_pos;\n      --len;\n    }\n    nbytes = br.bit_end_pos_ - br.bit_pos_ >> 3;\n    if (br_pos + nbytes > bit_reader_default.IBUF_MASK) {\n      var tail = bit_reader_default.IBUF_MASK + 1 - br_pos;\n      for (var x = 0; x < tail; x++)\n        ringbuffer[rb_pos + x] = br.buf_[br_pos + x];\n      nbytes -= tail;\n      rb_pos += tail;\n      len -= tail;\n      br_pos = 0;\n    }\n    for (var x = 0; x < nbytes; x++)\n      ringbuffer[rb_pos + x] = br.buf_[br_pos + x];\n    rb_pos += nbytes;\n    len -= nbytes;\n    if (rb_pos >= rb_size) {\n      output.write(ringbuffer, rb_size);\n      rb_pos -= rb_size;\n      for (var x = 0; x < rb_pos; x++)\n        ringbuffer[x] = ringbuffer[rb_size + x];\n    }\n    while (rb_pos + len >= rb_size) {\n      nbytes = rb_size - rb_pos;\n      if (br.input_.read(ringbuffer, rb_pos, nbytes) < nbytes) {\n        throw new Error("[CopyUncompressedBlockToOutput] not enough bytes");\n      }\n      output.write(ringbuffer, rb_size);\n      len -= nbytes;\n      rb_pos = 0;\n    }\n    if (br.input_.read(ringbuffer, rb_pos, len) < len) {\n      throw new Error("[CopyUncompressedBlockToOutput] not enough bytes");\n    }\n    br.reset();\n  }\n  function JumpToByteBoundary(br) {\n    var new_bit_pos = br.bit_pos_ + 7 & ~7;\n    var pad_bits = br.readBits(new_bit_pos - br.bit_pos_);\n    return pad_bits == 0;\n  }\n  function BrotliDecompressedSize(buffer) {\n    var input = new BrotliInput(buffer);\n    var br = new bit_reader_default(input);\n    DecodeWindowBits(br);\n    var out = DecodeMetaBlockLength(br);\n    return out.meta_block_length;\n  }\n  function BrotliDecompressBuffer(buffer, output_size) {\n    var input = new BrotliInput(buffer);\n    if (output_size == null) {\n      output_size = BrotliDecompressedSize(buffer);\n    }\n    var output_buffer = new Uint8Array(output_size);\n    var output = new BrotliOutput(output_buffer);\n    BrotliDecompress(input, output);\n    if (output.pos < output.buffer.length) {\n      output.buffer = output.buffer.subarray(0, output.pos);\n    }\n    return output.buffer;\n  }\n  function BrotliDecompress(input, output) {\n    var i;\n    var pos = 0;\n    var input_end = 0;\n    var window_bits = 0;\n    var max_backward_distance;\n    var max_distance = 0;\n    var ringbuffer_size;\n    var ringbuffer_mask;\n    var ringbuffer;\n    var ringbuffer_end;\n    var dist_rb = [16, 15, 11, 4];\n    var dist_rb_idx = 0;\n    var prev_byte1 = 0;\n    var prev_byte2 = 0;\n    var hgroup = [new HuffmanTreeGroup(0, 0), new HuffmanTreeGroup(0, 0), new HuffmanTreeGroup(0, 0)];\n    var block_type_trees;\n    var block_len_trees;\n    var br;\n    var kRingBufferWriteAheadSlack = 128 + bit_reader_default.READ_SIZE;\n    br = new bit_reader_default(input);\n    window_bits = DecodeWindowBits(br);\n    max_backward_distance = (1 << window_bits) - 16;\n    ringbuffer_size = 1 << window_bits;\n    ringbuffer_mask = ringbuffer_size - 1;\n    ringbuffer = new Uint8Array(ringbuffer_size + kRingBufferWriteAheadSlack + maxDictionaryWordLength);\n    ringbuffer_end = ringbuffer_size;\n    block_type_trees = [];\n    block_len_trees = [];\n    for (var x = 0; x < 3 * HUFFMAN_MAX_TABLE_SIZE; x++) {\n      block_type_trees[x] = new HuffmanCode(0, 0);\n      block_len_trees[x] = new HuffmanCode(0, 0);\n    }\n    while (!input_end) {\n      var meta_block_remaining_len = 0;\n      var is_uncompressed;\n      var block_length = [1 << 28, 1 << 28, 1 << 28];\n      var block_type = [0];\n      var num_block_types = [1, 1, 1];\n      var block_type_rb = [0, 1, 0, 1, 0, 1];\n      var block_type_rb_index = [0];\n      var distance_postfix_bits;\n      var num_direct_distance_codes;\n      var distance_postfix_mask;\n      var num_distance_codes;\n      var context_map = null;\n      var context_modes = null;\n      var num_literal_htrees;\n      var dist_context_map = null;\n      var num_dist_htrees;\n      var context_offset = 0;\n      var context_map_slice = null;\n      var literal_htree_index = 0;\n      var dist_context_offset = 0;\n      var dist_context_map_slice = null;\n      var dist_htree_index = 0;\n      var context_lookup_offset1 = 0;\n      var context_lookup_offset2 = 0;\n      var context_mode;\n      var htree_command;\n      for (i = 0; i < 3; ++i) {\n        hgroup[i].codes = null;\n        hgroup[i].htrees = null;\n      }\n      br.readMoreInput();\n      var _out = DecodeMetaBlockLength(br);\n      meta_block_remaining_len = _out.meta_block_length;\n      if (pos + meta_block_remaining_len > output.buffer.length) {\n        var tmp = new Uint8Array(pos + meta_block_remaining_len);\n        tmp.set(output.buffer);\n        output.buffer = tmp;\n      }\n      input_end = _out.input_end;\n      is_uncompressed = _out.is_uncompressed;\n      if (_out.is_metadata) {\n        JumpToByteBoundary(br);\n        for (; meta_block_remaining_len > 0; --meta_block_remaining_len) {\n          br.readMoreInput();\n          br.readBits(8);\n        }\n        continue;\n      }\n      if (meta_block_remaining_len === 0) {\n        continue;\n      }\n      if (is_uncompressed) {\n        br.bit_pos_ = br.bit_pos_ + 7 & ~7;\n        CopyUncompressedBlockToOutput(\n          output,\n          meta_block_remaining_len,\n          pos,\n          ringbuffer,\n          ringbuffer_mask,\n          br\n        );\n        pos += meta_block_remaining_len;\n        continue;\n      }\n      for (i = 0; i < 3; ++i) {\n        num_block_types[i] = DecodeVarLenUint8(br) + 1;\n        if (num_block_types[i] >= 2) {\n          ReadHuffmanCode(num_block_types[i] + 2, block_type_trees, i * HUFFMAN_MAX_TABLE_SIZE, br);\n          ReadHuffmanCode(kNumBlockLengthCodes, block_len_trees, i * HUFFMAN_MAX_TABLE_SIZE, br);\n          block_length[i] = ReadBlockLength(block_len_trees, i * HUFFMAN_MAX_TABLE_SIZE, br);\n          block_type_rb_index[i] = 1;\n        }\n      }\n      br.readMoreInput();\n      distance_postfix_bits = br.readBits(2);\n      num_direct_distance_codes = NUM_DISTANCE_SHORT_CODES + (br.readBits(4) << distance_postfix_bits);\n      distance_postfix_mask = (1 << distance_postfix_bits) - 1;\n      num_distance_codes = num_direct_distance_codes + (48 << distance_postfix_bits);\n      context_modes = new Uint8Array(num_block_types[0]);\n      for (i = 0; i < num_block_types[0]; ++i) {\n        br.readMoreInput();\n        context_modes[i] = br.readBits(2) << 1;\n      }\n      var _o1 = DecodeContextMap(num_block_types[0] << kLiteralContextBits, br);\n      num_literal_htrees = _o1.num_htrees;\n      context_map = _o1.context_map;\n      var _o2 = DecodeContextMap(num_block_types[2] << kDistanceContextBits, br);\n      num_dist_htrees = _o2.num_htrees;\n      dist_context_map = _o2.context_map;\n      hgroup[0] = new HuffmanTreeGroup(kNumLiteralCodes, num_literal_htrees);\n      hgroup[1] = new HuffmanTreeGroup(kNumInsertAndCopyCodes, num_block_types[1]);\n      hgroup[2] = new HuffmanTreeGroup(num_distance_codes, num_dist_htrees);\n      for (i = 0; i < 3; ++i) {\n        hgroup[i].decode(br);\n      }\n      context_map_slice = 0;\n      dist_context_map_slice = 0;\n      context_mode = context_modes[block_type[0]];\n      context_lookup_offset1 = lookupOffsets[context_mode];\n      context_lookup_offset2 = lookupOffsets[context_mode + 1];\n      htree_command = hgroup[1].htrees[0];\n      while (meta_block_remaining_len > 0) {\n        var cmd_code;\n        var range_idx;\n        var insert_code;\n        var copy_code;\n        var insert_length;\n        var copy_length;\n        var distance_code;\n        var distance;\n        var context;\n        var j;\n        var copy_dst;\n        br.readMoreInput();\n        if (block_length[1] === 0) {\n          DecodeBlockType(\n            num_block_types[1],\n            block_type_trees,\n            1,\n            block_type,\n            block_type_rb,\n            block_type_rb_index,\n            br\n          );\n          block_length[1] = ReadBlockLength(block_len_trees, HUFFMAN_MAX_TABLE_SIZE, br);\n          htree_command = hgroup[1].htrees[block_type[1]];\n        }\n        --block_length[1];\n        cmd_code = ReadSymbol(hgroup[1].codes, htree_command, br);\n        range_idx = cmd_code >> 6;\n        if (range_idx >= 2) {\n          range_idx -= 2;\n          distance_code = -1;\n        } else {\n          distance_code = 0;\n        }\n        insert_code = kInsertRangeLut[range_idx] + (cmd_code >> 3 & 7);\n        copy_code = kCopyRangeLut[range_idx] + (cmd_code & 7);\n        insert_length = kInsertLengthPrefixCode[insert_code].offset + br.readBits(kInsertLengthPrefixCode[insert_code].nbits);\n        copy_length = kCopyLengthPrefixCode[copy_code].offset + br.readBits(kCopyLengthPrefixCode[copy_code].nbits);\n        prev_byte1 = ringbuffer[pos - 1 & ringbuffer_mask];\n        prev_byte2 = ringbuffer[pos - 2 & ringbuffer_mask];\n        for (j = 0; j < insert_length; ++j) {\n          br.readMoreInput();\n          if (block_length[0] === 0) {\n            DecodeBlockType(\n              num_block_types[0],\n              block_type_trees,\n              0,\n              block_type,\n              block_type_rb,\n              block_type_rb_index,\n              br\n            );\n            block_length[0] = ReadBlockLength(block_len_trees, 0, br);\n            context_offset = block_type[0] << kLiteralContextBits;\n            context_map_slice = context_offset;\n            context_mode = context_modes[block_type[0]];\n            context_lookup_offset1 = lookupOffsets[context_mode];\n            context_lookup_offset2 = lookupOffsets[context_mode + 1];\n          }\n          context = lookup[context_lookup_offset1 + prev_byte1] | lookup[context_lookup_offset2 + prev_byte2];\n          literal_htree_index = context_map[context_map_slice + context];\n          --block_length[0];\n          prev_byte2 = prev_byte1;\n          prev_byte1 = ReadSymbol(hgroup[0].codes, hgroup[0].htrees[literal_htree_index], br);\n          ringbuffer[pos & ringbuffer_mask] = prev_byte1;\n          if ((pos & ringbuffer_mask) === ringbuffer_mask) {\n            output.write(ringbuffer, ringbuffer_size);\n          }\n          ++pos;\n        }\n        meta_block_remaining_len -= insert_length;\n        if (meta_block_remaining_len <= 0) break;\n        if (distance_code < 0) {\n          var context;\n          br.readMoreInput();\n          if (block_length[2] === 0) {\n            DecodeBlockType(\n              num_block_types[2],\n              block_type_trees,\n              2,\n              block_type,\n              block_type_rb,\n              block_type_rb_index,\n              br\n            );\n            block_length[2] = ReadBlockLength(block_len_trees, 2 * HUFFMAN_MAX_TABLE_SIZE, br);\n            dist_context_offset = block_type[2] << kDistanceContextBits;\n            dist_context_map_slice = dist_context_offset;\n          }\n          --block_length[2];\n          context = (copy_length > 4 ? 3 : copy_length - 2) & 255;\n          dist_htree_index = dist_context_map[dist_context_map_slice + context];\n          distance_code = ReadSymbol(hgroup[2].codes, hgroup[2].htrees[dist_htree_index], br);\n          if (distance_code >= num_direct_distance_codes) {\n            var nbits;\n            var postfix;\n            var offset;\n            distance_code -= num_direct_distance_codes;\n            postfix = distance_code & distance_postfix_mask;\n            distance_code >>= distance_postfix_bits;\n            nbits = (distance_code >> 1) + 1;\n            offset = (2 + (distance_code & 1) << nbits) - 4;\n            distance_code = num_direct_distance_codes + (offset + br.readBits(nbits) << distance_postfix_bits) + postfix;\n          }\n        }\n        distance = TranslateShortCodes(distance_code, dist_rb, dist_rb_idx);\n        if (distance < 0) {\n          throw new Error("[BrotliDecompress] invalid distance");\n        }\n        if (pos < max_backward_distance && max_distance !== max_backward_distance) {\n          max_distance = pos;\n        } else {\n          max_distance = max_backward_distance;\n        }\n        copy_dst = pos & ringbuffer_mask;\n        if (distance > max_distance) {\n          if (copy_length >= minDictionaryWordLength && copy_length <= maxDictionaryWordLength) {\n            var dictOffset = offsetsByLength[copy_length];\n            var word_id = distance - max_distance - 1;\n            var shift = sizeBitsByLength[copy_length];\n            var mask = (1 << shift) - 1;\n            var word_idx = word_id & mask;\n            var transform_idx = word_id >> shift;\n            dictOffset += word_idx * copy_length;\n            if (transform_idx < kNumTransforms) {\n              var len = transformDictionaryWord(ringbuffer, copy_dst, dictOffset, copy_length, transform_idx);\n              copy_dst += len;\n              pos += len;\n              meta_block_remaining_len -= len;\n              if (copy_dst >= ringbuffer_end) {\n                output.write(ringbuffer, ringbuffer_size);\n                for (var _x = 0; _x < copy_dst - ringbuffer_end; _x++)\n                  ringbuffer[_x] = ringbuffer[ringbuffer_end + _x];\n              }\n            } else {\n              throw new Error("Invalid backward reference. pos: " + pos + " distance: " + distance + " len: " + copy_length + " bytes left: " + meta_block_remaining_len);\n            }\n          } else {\n            throw new Error("Invalid backward reference. pos: " + pos + " distance: " + distance + " len: " + copy_length + " bytes left: " + meta_block_remaining_len);\n          }\n        } else {\n          if (distance_code > 0) {\n            dist_rb[dist_rb_idx & 3] = distance;\n            ++dist_rb_idx;\n          }\n          if (copy_length > meta_block_remaining_len) {\n            throw new Error("Invalid backward reference. pos: " + pos + " distance: " + distance + " len: " + copy_length + " bytes left: " + meta_block_remaining_len);\n          }\n          for (j = 0; j < copy_length; ++j) {\n            ringbuffer[pos & ringbuffer_mask] = ringbuffer[pos - distance & ringbuffer_mask];\n            if ((pos & ringbuffer_mask) === ringbuffer_mask) {\n              output.write(ringbuffer, ringbuffer_size);\n            }\n            ++pos;\n            --meta_block_remaining_len;\n          }\n        }\n        prev_byte1 = ringbuffer[pos - 1 & ringbuffer_mask];\n        prev_byte2 = ringbuffer[pos - 2 & ringbuffer_mask];\n      }\n      pos &= 1073741823;\n    }\n    output.write(ringbuffer, pos & ringbuffer_mask);\n  }\n  var kDefaultCodeLength, kCodeLengthRepeatCode, kNumLiteralCodes, kNumInsertAndCopyCodes, kNumBlockLengthCodes, kLiteralContextBits, kDistanceContextBits, HUFFMAN_TABLE_BITS, HUFFMAN_TABLE_MASK, HUFFMAN_MAX_TABLE_SIZE, CODE_LENGTH_CODES, kCodeLengthCodeOrder, NUM_DISTANCE_SHORT_CODES, kDistanceShortCodeIndexOffset, kDistanceShortCodeValueOffset, kMaxHuffmanTableSize;\n  var init_decode = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/dec/decode.ts"() {\n      init_streams();\n      init_bit_reader();\n      init_dictionary();\n      init_huffman();\n      init_context();\n      init_prefix();\n      init_transform();\n      kDefaultCodeLength = 8;\n      kCodeLengthRepeatCode = 16;\n      kNumLiteralCodes = 256;\n      kNumInsertAndCopyCodes = 704;\n      kNumBlockLengthCodes = 26;\n      kLiteralContextBits = 6;\n      kDistanceContextBits = 2;\n      HUFFMAN_TABLE_BITS = 8;\n      HUFFMAN_TABLE_MASK = 255;\n      HUFFMAN_MAX_TABLE_SIZE = 1080;\n      CODE_LENGTH_CODES = 18;\n      kCodeLengthCodeOrder = new Uint8Array([\n        1,\n        2,\n        3,\n        4,\n        0,\n        5,\n        17,\n        6,\n        16,\n        7,\n        8,\n        9,\n        10,\n        11,\n        12,\n        13,\n        14,\n        15\n      ]);\n      NUM_DISTANCE_SHORT_CODES = 16;\n      kDistanceShortCodeIndexOffset = new Uint8Array([\n        3,\n        2,\n        1,\n        0,\n        3,\n        3,\n        3,\n        3,\n        3,\n        3,\n        2,\n        2,\n        2,\n        2,\n        2,\n        2\n      ]);\n      kDistanceShortCodeValueOffset = new Int8Array([\n        0,\n        0,\n        0,\n        0,\n        -1,\n        1,\n        -2,\n        2,\n        -3,\n        3,\n        -1,\n        1,\n        -2,\n        2,\n        -3,\n        3\n      ]);\n      kMaxHuffmanTableSize = new Uint16Array([\n        256,\n        402,\n        436,\n        468,\n        500,\n        534,\n        566,\n        598,\n        630,\n        662,\n        694,\n        726,\n        758,\n        790,\n        822,\n        854,\n        886,\n        920,\n        952,\n        984,\n        1016,\n        1048,\n        1080\n      ]);\n      HuffmanTreeGroup.prototype.decode = function(br) {\n        var i;\n        var table_size;\n        var next = 0;\n        for (i = 0; i < this.num_htrees; ++i) {\n          this.htrees[i] = next;\n          table_size = ReadHuffmanCode(this.alphabet_size, this.codes, next, br);\n          next += table_size;\n        }\n      };\n      init2();\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/index.ts\n  var brotli_dec_exports = {};\n  __export(brotli_dec_exports, {\n    brotliDecompress: () => brotliDecompress\n  });\n  function brotliDecompress(input) {\n    return BrotliDecompressBuffer(input);\n  }\n  var init_brotli_dec = __esm({\n    "vendor/onlyoffice-web-comp/internal/vendor/brotli-dec/index.ts"() {\n      init_decode();\n    }\n  });\n\n  // vendor/onlyoffice-web-comp/const/index.ts\n  var OFFICE_THEME = {\n    LIGHT: "theme-light",\n    CLASSIC_LIGHT: "theme-classic-light",\n    WHITE: "theme-white",\n    DARK: "theme-dark",\n    NIGHT: "theme-night",\n    CONTRAST_DARK: "theme-contrast-dark"\n  };\n  var DEFAULT_OFFICE_THEME = OFFICE_THEME.WHITE;\n  var OFFICE_THEME_OPTIONS = [\n    { id: OFFICE_THEME.WHITE, label: "浅色" },\n    { id: OFFICE_THEME.CLASSIC_LIGHT, label: "经典浅色" },\n    { id: OFFICE_THEME.LIGHT, label: "Light" },\n    { id: OFFICE_THEME.DARK, label: "深色" },\n    { id: OFFICE_THEME.NIGHT, label: "夜间" },\n    { id: OFFICE_THEME.CONTRAST_DARK, label: "高对比深色" }\n  ];\n  var OFFICE_XML_EVENT_CONFIG = {\n    default: {\n      isEnable: false,\n      limitBytes: 1024 * 1024 * 2048\n    }\n  };\n  var X2T_PDF_DEFAULT_FONT_FILE = "Carlito-Regular.ttf";\n  var X2T_PDF_FONT_MANIFEST = [\n    {\n      file: "Carlito-Regular.ttf",\n      aliases: ["Carlito.ttf", "Calibri.ttf"]\n    },\n    {\n      file: "Carlito-Bold.ttf",\n      aliases: ["Carlito_Bold.ttf", "Calibri_Bold.ttf"]\n    },\n    {\n      file: "Carlito-Italic.ttf",\n      aliases: ["Carlito_Italic.ttf", "Calibri_Italic.ttf"]\n    },\n    {\n      file: "Carlito-BoldItalic.ttf",\n      aliases: ["Carlito_Bold_Italic.ttf", "Calibri_Bold_Italic.ttf"]\n    },\n    {\n      file: "Arial-Regular.ttf",\n      aliases: ["Arial.ttf"]\n    },\n    {\n      file: "Arial-Bold.ttf",\n      aliases: ["Arial_Bold.ttf"]\n    },\n    {\n      file: "Arial-Italic.ttf",\n      aliases: ["Arial_Italic.ttf"]\n    },\n    {\n      file: "Arial-BoldItalic.ttf",\n      aliases: ["Arial_Bold_Italic.ttf"]\n    },\n    {\n      file: "DroidSansFallback.ttf",\n      aliases: [\n        "Droid Sans Fallback.ttf",\n        "SimSun.ttf",\n        "NSimSun.ttf",\n        "宋体.ttf",\n        "Microsoft YaHei.ttf",\n        "微软雅黑.ttf",\n        "PingFang SC.ttf"\n      ]\n    }\n  ];\n  var DEFAULT_ONLYOFFICE_VERSION = "9.4.0-develop";\n  var DEFAULT_ONLYOFFICE_ROOT = `/packages/onlyoffice/${DEFAULT_ONLYOFFICE_VERSION}`;\n  var staticResourceOptions = null;\n  function trimTrailingSlash(value) {\n    return value.replace(/\\/+$/, "");\n  }\n  function buildStaticResource() {\n    const apiJs = "/web-apps/apps/api/documents/api.js";\n    const preloadHtml = "/web-apps/apps/api/documents/preload.html";\n    const cdnOrigin = staticResourceOptions?.cdnOrigin ? trimTrailingSlash(staticResourceOptions.cdnOrigin) : "";\n    const onlyofficeVersion = staticResourceOptions?.onlyofficeVersion || DEFAULT_ONLYOFFICE_VERSION;\n    const onlyofficeRoot = cdnOrigin ? `${cdnOrigin}/onlyoffice/${onlyofficeVersion}` : DEFAULT_ONLYOFFICE_ROOT;\n    const x2tRoot = `${onlyofficeRoot}/x2t`;\n    const x2tPdfFontsRoot = `${onlyofficeRoot}/x2t-fonts`;\n    return {\n      version: {\n        onlyofficeSdk: onlyofficeRoot,\n        x2t: x2tRoot\n      },\n      onlyoffice: {\n        root: onlyofficeRoot,\n        apiJs,\n        preloadHtml,\n        apiUrl: onlyofficeRoot + apiJs,\n        preloadUrl: onlyofficeRoot + preloadHtml\n      },\n      x2t: {\n        root: x2tRoot,\n        script: `${x2tRoot}/x2t.js`,\n        wasm: `${x2tRoot}/x2t.wasm`,\n        pdfFonts: {\n          root: x2tPdfFontsRoot,\n          default: `${x2tPdfFontsRoot}/${X2T_PDF_DEFAULT_FONT_FILE}`\n        }\n      }\n    };\n  }\n  var staticResourceCache = null;\n  function getStaticResource() {\n    if (!staticResourceCache) {\n      staticResourceCache = buildStaticResource();\n    }\n    return staticResourceCache;\n  }\n  var STATIC_RESOURCE = {\n    get version() {\n      return getStaticResource().version;\n    },\n    get onlyoffice() {\n      return getStaticResource().onlyoffice;\n    },\n    get x2t() {\n      return getStaticResource().x2t;\n    }\n  };\n  function resolveSiteUrl(origin, path) {\n    if (/^[a-z][a-z0-9+.-]*:\\/\\//i.test(path)) {\n      return path;\n    }\n    return `${origin}${path.startsWith("/") ? path : `/${path}`}`;\n  }\n  var AppType = {\n    word: 1,\n    slide: 3,\n    cell: 2,\n    draw: 4,\n    pdf: 5\n  };\n  var docTypeMap = {\n    docx: AppType.word,\n    doc: AppType.word,\n    odt: AppType.word,\n    rtf: AppType.word,\n    txt: AppType.word,\n    html: AppType.word,\n    mht: AppType.word,\n    epub: AppType.word,\n    fb2: AppType.word,\n    mobi: AppType.word,\n    docm: AppType.word,\n    dotx: AppType.word,\n    dotm: AppType.word,\n    oform: AppType.word,\n    docxf: AppType.word,\n    pptx: AppType.slide,\n    ppt: AppType.slide,\n    odp: AppType.slide,\n    ppsx: AppType.slide,\n    pptm: AppType.slide,\n    ppsm: AppType.slide,\n    potx: AppType.slide,\n    potm: AppType.slide,\n    otp: AppType.slide,\n    odg: AppType.slide,\n    xlsx: AppType.cell,\n    xls: AppType.cell,\n    ods: AppType.cell,\n    csv: AppType.cell,\n    xlsm: AppType.cell,\n    xltx: AppType.cell,\n    xltm: AppType.cell,\n    xlsb: AppType.cell,\n    ots: AppType.cell,\n    vsdx: AppType.draw,\n    vssx: AppType.draw,\n    vstx: AppType.draw,\n    vsdm: AppType.draw,\n    vssm: AppType.draw,\n    vstm: AppType.draw,\n    pdf: AppType.pdf\n  };\n  var appTypeName = {\n    [AppType.word]: "word" /* Word */,\n    [AppType.slide]: "slide" /* Slide */,\n    [AppType.cell]: "cell" /* Cell */,\n    [AppType.draw]: "draw" /* Draw */,\n    [AppType.pdf]: "pdf" /* Pdf */\n  };\n\n  // vendor/onlyoffice-web-comp/internal/editor/x2t-assets.ts\n  var WASM_MAGIC = [0, 97, 115, 109];\n  var brotliStreamSupported = null;\n  function isWasmBinary(data) {\n    return data.length >= WASM_MAGIC.length && WASM_MAGIC.every((byte, index) => data[index] === byte);\n  }\n  function isJavaScriptSource(data) {\n    const head = new TextDecoder().decode(\n      data.slice(0, Math.min(data.length, 128))\n    );\n    return /^\\s*(\\/\\/|\\/\\*|var |function|\\(function)/.test(head);\n  }\n  function isAlreadyDecompressed(data) {\n    return isWasmBinary(data) || isJavaScriptSource(data);\n  }\n  function canDecompressBrotliInBrowser() {\n    if (brotliStreamSupported !== null) {\n      return brotliStreamSupported;\n    }\n    if (!("DecompressionStream" in globalThis)) {\n      brotliStreamSupported = false;\n      return false;\n    }\n    try {\n      new DecompressionStream("br");\n      brotliStreamSupported = true;\n    } catch {\n      brotliStreamSupported = false;\n    }\n    return brotliStreamSupported;\n  }\n  function toArrayBuffer(data) {\n    return data.buffer.slice(\n      data.byteOffset,\n      data.byteOffset + data.byteLength\n    );\n  }\n  async function decompressWithStream(buffer) {\n    const stream = new Blob([buffer]).stream().pipeThrough(new DecompressionStream("br"));\n    return new Response(stream).arrayBuffer();\n  }\n  function decompressWithVendoredBrotli(buffer) {\n    return Promise.resolve().then(() => (init_brotli_dec(), brotli_dec_exports)).then(\n      ({ brotliDecompress: brotliDecompress2 }) => toArrayBuffer(brotliDecompress2(new Uint8Array(buffer)))\n    );\n  }\n  async function decompressBrotli(buffer) {\n    if (canDecompressBrotliInBrowser()) {\n      try {\n        return await decompressWithStream(buffer);\n      } catch {\n      }\n    }\n    return decompressWithVendoredBrotli(buffer);\n  }\n  async function fetchMaybeBrotliAsset(url, fetchImpl = fetch) {\n    const response = await fetchImpl(url);\n    if (!response.ok) {\n      throw new Error(`Failed to fetch ${url}: ${response.status}`);\n    }\n    const buffer = await response.arrayBuffer();\n    const bytes = new Uint8Array(buffer);\n    if (isAlreadyDecompressed(bytes)) {\n      return buffer;\n    }\n    return decompressBrotli(buffer);\n  }\n  async function fetchMaybeBrotliScript(url, fetchImpl = fetch) {\n    const buffer = await fetchMaybeBrotliAsset(url, fetchImpl);\n    return new TextDecoder().decode(buffer);\n  }\n  var cachedFonts = null;\n  var cachedFontsRoot = "";\n  var loadingPromise = null;\n  var loadingFontsRoot = "";\n  async function fetchPdfFontFile(origin, root, file) {\n    const url = resolveSiteUrl(\n      origin,\n      `${root}/${file}`\n    );\n    const response = await fetch(url);\n    if (!response.ok) {\n      console.warn("[x2t-assets] font fetch failed:", url, response.status);\n      return null;\n    }\n    const bytes = new Uint8Array(await response.arrayBuffer());\n    return bytes.byteLength ? bytes : null;\n  }\n  async function loadX2tPdfFonts(origin, root = STATIC_RESOURCE.x2t.pdfFonts.root) {\n    if (cachedFonts && cachedFontsRoot === root) {\n      return cachedFonts;\n    }\n    if (loadingPromise && loadingFontsRoot === root) {\n      return loadingPromise;\n    }\n    loadingFontsRoot = root;\n    loadingPromise = (async () => {\n      const fonts = {};\n      for (const entry of X2T_PDF_FONT_MANIFEST) {\n        const bytes = await fetchPdfFontFile(origin, root, entry.file);\n        if (!bytes) {\n          continue;\n        }\n        for (const alias of entry.aliases) {\n          fonts[alias] = bytes;\n        }\n      }\n      if (!fonts["Carlito.ttf"]?.byteLength) {\n        console.warn(\n          "[x2t-assets] Carlito regular missing under",\n          root\n        );\n        return {};\n      }\n      cachedFonts = fonts;\n      cachedFontsRoot = root;\n      return cachedFonts;\n    })();\n    try {\n      return await loadingPromise;\n    } finally {\n      loadingPromise = null;\n    }\n  }\n\n  // vendor/onlyoffice-web-comp/internal/editor/x2t.worker.ts\n  var X2T_ORIGIN = self.location.origin;\n  var x2t = null;\n  var initPromise = null;\n  var initResourceKey = "";\n  var transparentPng = Uint8Array.from(\n    atob(\n      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="\n    ),\n    (value) => value.charCodeAt(0)\n  );\n  var crcTable = (() => {\n    const table = new Uint32Array(256);\n    for (let i = 0; i < 256; i++) {\n      let c = i;\n      for (let k = 0; k < 8; k++) {\n        c = c & 1 ? 3988292384 ^ c >>> 1 : c >>> 1;\n      }\n      table[i] = c >>> 0;\n    }\n    return table;\n  })();\n  function crc32(data) {\n    let crc = 4294967295;\n    for (const byte of data) {\n      crc = crcTable[(crc ^ byte) & 255] ^ crc >>> 8;\n    }\n    return (crc ^ 4294967295) >>> 0;\n  }\n  function readU16(data, offset) {\n    return data[offset] | data[offset + 1] << 8;\n  }\n  function readU32(data, offset) {\n    return (data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24) >>> 0;\n  }\n  function writeU16(data, offset, value) {\n    data[offset] = value & 255;\n    data[offset + 1] = value >>> 8 & 255;\n  }\n  function writeU32(data, offset, value) {\n    data[offset] = value & 255;\n    data[offset + 1] = value >>> 8 & 255;\n    data[offset + 2] = value >>> 16 & 255;\n    data[offset + 3] = value >>> 24 & 255;\n  }\n  function concatBytes(parts) {\n    const size = parts.reduce((total, part) => total + part.length, 0);\n    const output = new Uint8Array(size);\n    let offset = 0;\n    for (const part of parts) {\n      output.set(part, offset);\n      offset += part.length;\n    }\n    return output;\n  }\n  async function decompressDeflateRaw(data) {\n    if (!("DecompressionStream" in self)) {\n      return null;\n    }\n    const stream = new Blob([data]).stream().pipeThrough(new self.DecompressionStream("deflate-raw"));\n    return new Uint8Array(await new Response(stream).arrayBuffer());\n  }\n  function shouldRewriteZipEntry(name) {\n    return /\\.(xml|rels)$/i.test(name) && (name.startsWith("word/") || name.startsWith("xl/") || name.startsWith("ppt/"));\n  }\n  function shouldRewriteDocxBookmarkEntry(name) {\n    return /^word\\/.+\\.xml$/i.test(name);\n  }\n  function rewriteFontText(text, aliases) {\n    if (!aliases) return text;\n    const entries = Object.entries(aliases).filter(([from, to]) => from && to && from !== to).sort(([a], [b]) => b.length - a.length);\n    let output = text;\n    for (const [from, to] of entries) {\n      const escaped = from.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&");\n      const suffix = to.startsWith(from) && to.length > from.length ? to.slice(from.length).replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&") : "";\n      const pattern = suffix ? `${escaped}(?!${suffix})` : escaped;\n      output = output.replace(new RegExp(pattern, "gi"), to);\n    }\n    return output;\n  }\n  function hasOfficeZipExtension(fileName) {\n    return /\\.(docx|xlsx|pptx)$/i.test(fileName);\n  }\n  function hasDocxExtension(fileName) {\n    return /\\.docx$/i.test(fileName);\n  }\n  function isDecimalBookmarkId(value) {\n    return /^\\d+$/.test(value);\n  }\n  function normalizeDocxBookmarkIds(text) {\n    const bookmarkTagPattern = /<w:bookmark(Start|End)\\b[^>]*\\bw:id="([^"]*)"[^>]*>/g;\n    const usedNumericIds = /* @__PURE__ */ new Set();\n    let maxNumericId = -1;\n    let hasInvalidId = false;\n    text.replace(bookmarkTagPattern, (_tag, _kind, id) => {\n      if (isDecimalBookmarkId(id)) {\n        usedNumericIds.add(id);\n        maxNumericId = Math.max(maxNumericId, Number(id));\n      } else {\n        hasInvalidId = true;\n      }\n      return _tag;\n    });\n    if (!hasInvalidId) return text;\n    const remappedIds = /* @__PURE__ */ new Map();\n    const nextId = () => {\n      do {\n        maxNumericId += 1;\n      } while (usedNumericIds.has(String(maxNumericId)));\n      const value = String(maxNumericId);\n      usedNumericIds.add(value);\n      return value;\n    };\n    return text.replace(bookmarkTagPattern, (tag, _kind, id) => {\n      if (isDecimalBookmarkId(id)) return tag;\n      let mapped = remappedIds.get(id);\n      if (!mapped) {\n        mapped = nextId();\n        remappedIds.set(id, mapped);\n      }\n      return tag.replace(/\\bw:id="[^"]*"/, `w:id="${mapped}"`);\n    });\n  }\n  function encodeUtf16Le(value) {\n    const output = new Uint8Array(value.length * 2);\n    for (let i = 0; i < value.length; i++) {\n      const code = value.charCodeAt(i);\n      output[i * 2] = code & 255;\n      output[i * 2 + 1] = code >>> 8;\n    }\n    return output;\n  }\n  function replaceBytesWithPadding(data, from, to, skipSuffix) {\n    if (!from.length || to.length > from.length) return data;\n    const output = data.slice();\n    for (let i = 0; i <= output.length - from.length; i++) {\n      let matched = true;\n      for (let j = 0; j < from.length; j++) {\n        if (output[i + j] !== from[j]) {\n          matched = false;\n          break;\n        }\n      }\n      if (!matched) continue;\n      if (skipSuffix && skipSuffix.length) {\n        let suffixMatched = true;\n        for (let k = 0; k < skipSuffix.length; k++) {\n          if (output[i + from.length + k] !== skipSuffix[k]) {\n            suffixMatched = false;\n            break;\n          }\n        }\n        if (suffixMatched) continue;\n      }\n      output.fill(0, i, i + from.length);\n      output.set(to, i);\n      i += from.length - 1;\n    }\n    return output;\n  }\n  function restoreEditorBinFontNames(data, restoreAliases) {\n    if (!restoreAliases) return data;\n    const entries = Object.entries(restoreAliases).sort(\n      (a, b) => b[0].length - a[0].length\n    );\n    let output = data;\n    for (const [from, to] of entries) {\n      const skipSuffix = to.startsWith(from) && to.length > from.length ? encodeUtf16Le(to.slice(from.length)) : void 0;\n      output = replaceBytesWithPadding(\n        output,\n        encodeUtf16Le(from),\n        encodeUtf16Le(to),\n        skipSuffix\n      );\n    }\n    return output;\n  }\n  async function rewriteOfficeZipText(data, shouldRewriteEntry, rewriteText) {\n    if (!("TextDecoder" in self) || !("TextEncoder" in self)) {\n      return data;\n    }\n    const input = new Uint8Array(data);\n    let eocd = -1;\n    for (let i = input.length - 22; i >= Math.max(0, input.length - 65558); i--) {\n      if (readU32(input, i) === 101010256) {\n        eocd = i;\n        break;\n      }\n    }\n    if (eocd < 0) return data;\n    const entryCount = readU16(input, eocd + 10);\n    const centralOffset = readU32(input, eocd + 16);\n    const decoder = new TextDecoder();\n    const encoder = new TextEncoder();\n    const localParts = [];\n    const centralParts = [];\n    let readOffset = centralOffset;\n    let outputOffset = 0;\n    let changed = false;\n    for (let i = 0; i < entryCount; i++) {\n      if (readU32(input, readOffset) !== 33639248) return data;\n      const method = readU16(input, readOffset + 10);\n      const modTime = readU16(input, readOffset + 12);\n      const modDate = readU16(input, readOffset + 14);\n      const crc = readU32(input, readOffset + 16);\n      const compressedSize = readU32(input, readOffset + 20);\n      const uncompressedSize = readU32(input, readOffset + 24);\n      const nameLength = readU16(input, readOffset + 28);\n      const extraLength = readU16(input, readOffset + 30);\n      const commentLength = readU16(input, readOffset + 32);\n      const internalAttrs = readU16(input, readOffset + 36);\n      const externalAttrs = readU32(input, readOffset + 38);\n      const localOffset = readU32(input, readOffset + 42);\n      const nameBytes = input.slice(readOffset + 46, readOffset + 46 + nameLength);\n      const name = decoder.decode(nameBytes);\n      const localNameLength = readU16(input, localOffset + 26);\n      const localExtraLength = readU16(input, localOffset + 28);\n      const dataStart = localOffset + 30 + localNameLength + localExtraLength;\n      const compressed = input.slice(dataStart, dataStart + compressedSize);\n      let outputData = compressed;\n      let outputMethod = method;\n      let outputCrc = crc;\n      let outputUncompressedSize = uncompressedSize;\n      if (shouldRewriteEntry(name)) {\n        const uncompressed = method === 0 ? compressed : method === 8 ? await decompressDeflateRaw(compressed) : null;\n        if (uncompressed) {\n          const originalText = decoder.decode(uncompressed);\n          const rewrittenText = rewriteText(originalText, name);\n          if (rewrittenText !== originalText) {\n            const rewritten = encoder.encode(rewrittenText);\n            outputData = rewritten;\n            outputMethod = 0;\n            outputCrc = crc32(rewritten);\n            outputUncompressedSize = rewritten.length;\n            changed = true;\n          }\n        }\n      }\n      const localHeader = new Uint8Array(30 + nameBytes.length);\n      writeU32(localHeader, 0, 67324752);\n      writeU16(localHeader, 4, 20);\n      writeU16(localHeader, 6, 0);\n      writeU16(localHeader, 8, outputMethod);\n      writeU16(localHeader, 10, modTime);\n      writeU16(localHeader, 12, modDate);\n      writeU32(localHeader, 14, outputCrc);\n      writeU32(localHeader, 18, outputData.length);\n      writeU32(localHeader, 22, outputUncompressedSize);\n      writeU16(localHeader, 26, nameBytes.length);\n      writeU16(localHeader, 28, 0);\n      localHeader.set(nameBytes, 30);\n      localParts.push(localHeader, outputData);\n      const centralHeader = new Uint8Array(46 + nameBytes.length);\n      writeU32(centralHeader, 0, 33639248);\n      writeU16(centralHeader, 4, 20);\n      writeU16(centralHeader, 6, 20);\n      writeU16(centralHeader, 8, 0);\n      writeU16(centralHeader, 10, outputMethod);\n      writeU16(centralHeader, 12, modTime);\n      writeU16(centralHeader, 14, modDate);\n      writeU32(centralHeader, 16, outputCrc);\n      writeU32(centralHeader, 20, outputData.length);\n      writeU32(centralHeader, 24, outputUncompressedSize);\n      writeU16(centralHeader, 28, nameBytes.length);\n      writeU16(centralHeader, 30, 0);\n      writeU16(centralHeader, 32, 0);\n      writeU16(centralHeader, 34, 0);\n      writeU16(centralHeader, 36, internalAttrs);\n      writeU32(centralHeader, 38, externalAttrs);\n      writeU32(centralHeader, 42, outputOffset);\n      centralHeader.set(nameBytes, 46);\n      centralParts.push(centralHeader);\n      outputOffset += localHeader.length + outputData.length;\n      readOffset += 46 + nameLength + extraLength + commentLength;\n    }\n    const centralDirectory = concatBytes(centralParts);\n    const end = new Uint8Array(22);\n    writeU32(end, 0, 101010256);\n    writeU16(end, 8, entryCount);\n    writeU16(end, 10, entryCount);\n    writeU32(end, 12, centralDirectory.length);\n    writeU32(end, 16, outputOffset);\n    return changed ? concatBytes([...localParts, centralDirectory, end]).buffer : data;\n  }\n  async function rewriteZipFontNames(data, fileFrom, fontAliases) {\n    if (!fontAliases || !hasOfficeZipExtension(fileFrom)) {\n      return data;\n    }\n    return rewriteOfficeZipText(\n      data,\n      shouldRewriteZipEntry,\n      (text) => rewriteFontText(text, fontAliases)\n    );\n  }\n  async function sanitizeDocxForX2t(data, fileFrom) {\n    if (!hasDocxExtension(fileFrom)) {\n      return data;\n    }\n    return rewriteOfficeZipText(\n      data,\n      shouldRewriteDocxBookmarkEntry,\n      normalizeDocxBookmarkIds\n    );\n  }\n  function executeEmscriptenScript(scriptSource) {\n    if (typeof importScripts === "function") {\n      const scriptBlob = new Blob([scriptSource], {\n        type: "application/javascript"\n      });\n      const scriptBlobUrl = URL.createObjectURL(scriptBlob);\n      try {\n        importScripts(scriptBlobUrl);\n        return;\n      } catch (error) {\n        const isImportScriptsUnsupported = error instanceof TypeError && String(error.message).includes("importScripts");\n        if (!isImportScriptsUnsupported) {\n          throw error;\n        }\n      } finally {\n        URL.revokeObjectURL(scriptBlobUrl);\n      }\n    }\n    (0, eval)(scriptSource);\n  }\n  function getWorkerStaticResource(params) {\n    return params?.staticResource ?? getStaticResource();\n  }\n  async function initX2t(params) {\n    if (x2t) return;\n    const staticResource = getWorkerStaticResource(params);\n    const x2tBaseUrl = resolveSiteUrl(X2T_ORIGIN, `${staticResource.x2t.root}/`);\n    const scriptUrl = resolveSiteUrl(X2T_ORIGIN, staticResource.x2t.script);\n    const wasmUrl = resolveSiteUrl(X2T_ORIGIN, staticResource.x2t.wasm);\n    const [scriptSource, wasmBinary] = await Promise.all([\n      fetchMaybeBrotliScript(scriptUrl),\n      fetchMaybeBrotliAsset(wasmUrl)\n    ]);\n    Object.assign(self, {\n      __filename: x2tBaseUrl,\n      wasmBinary\n    });\n    executeEmscriptenScript(scriptSource);\n    x2t = self.Module;\n    await new Promise((resolve) => {\n      x2t.onRuntimeInitialized = () => resolve();\n    });\n    try {\n      x2t.FS.mkdir("/working");\n      x2t.FS.mkdir("/working/media");\n      x2t.FS.mkdir("/working/fonts");\n      x2t.FS.mkdir("/working/themes");\n    } catch (err) {\n      console.error("[x2t.worker] mkdir error:", err);\n    }\n    console.log("[x2t.worker] Initialized successfully");\n  }\n  async function ensureInit(params) {\n    const staticResource = getWorkerStaticResource(params);\n    const resourceKey = JSON.stringify(staticResource.x2t);\n    if (x2t && initResourceKey && initResourceKey !== resourceKey) {\n      throw new Error("x2t static resource changed after worker initialization");\n    }\n    initResourceKey = resourceKey;\n    if (!initPromise) {\n      initPromise = initX2t(params);\n    }\n    return initPromise;\n  }\n  function cleanupFiles(files) {\n    for (const file of files) {\n      try {\n        x2t.FS.unlink(file);\n      } catch (err) {\n        console.error(err);\n      }\n    }\n    cleanMedia();\n    cleanFonts();\n    cleanThemes();\n  }\n  var X2tWorkerError = class extends Error {\n    constructor(message, details) {\n      super(message);\n      this.details = details;\n      this.name = "X2tWorkerError";\n    }\n  };\n  function formatSerializedWorkerError(error) {\n    const message = error.message ? `: ${error.message}` : "";\n    const details = [\n      ["errno", error.errno],\n      ["code", error.code],\n      ["path", error.path]\n    ].filter(([, value]) => value != null).map(([key, value]) => `${key}=${String(value)}`).join(", ");\n    return details ? `${error.name}${message} (${details})` : `${error.name}${message}`;\n  }\n  function serializeWorkerError(error) {\n    if (error instanceof Error) {\n      const serialized = {\n        name: error.name || "Error",\n        message: error.message\n      };\n      return {\n        ...serialized,\n        text: formatSerializedWorkerError(serialized)\n      };\n    }\n    if (error && typeof error === "object") {\n      const record = error;\n      const constructorName = "constructor" in record && record.constructor && typeof record.constructor === "function" ? record.constructor.name : "";\n      const name = typeof record.name === "string" && record.name ? record.name : constructorName || Object.prototype.toString.call(error);\n      const message = typeof record.message === "string" && record.message ? record.message : void 0;\n      const serialized = {\n        name,\n        message,\n        errno: record.errno,\n        code: record.code,\n        path: record.path\n      };\n      return {\n        ...serialized,\n        text: formatSerializedWorkerError(serialized)\n      };\n    }\n    const text = String(error);\n    return {\n      name: "Error",\n      message: text,\n      text\n    };\n  }\n  function cleanMedia() {\n    try {\n      const mediaFiles = x2t.FS.readdir("/working/media/");\n      for (const file of mediaFiles) {\n        if (file !== "." && file !== "..") {\n          x2t.FS.unlink("/working/media/" + file);\n        }\n      }\n    } catch (err) {\n      console.error(err);\n    }\n  }\n  function cleanFonts() {\n    try {\n      const fontFiles = x2t.FS.readdir("/working/fonts/");\n      for (const file of fontFiles) {\n        if (file !== "." && file !== "..") {\n          x2t.FS.unlink("/working/fonts/" + file);\n        }\n      }\n    } catch (err) {\n      console.error(err);\n    }\n  }\n  function cleanThemes() {\n    cleanDirectoryFiles("/working/themes");\n  }\n  function cleanDirectoryFiles(dir) {\n    try {\n      const files = x2t.FS.readdir(dir);\n      for (const file of files) {\n        if (file === "." || file === "..") {\n          continue;\n        }\n        const path = `${dir}/${file}`;\n        const stat = x2t.FS.stat(path);\n        if (x2t.FS.isDir(stat.mode)) {\n          cleanDirectoryFiles(path);\n          x2t.FS.rmdir(path);\n        } else {\n          x2t.FS.unlink(path);\n        }\n      }\n    } catch (err) {\n      console.error(err);\n    }\n  }\n  function ensureDirectory(path) {\n    const parts = path.split("/").filter(Boolean);\n    let current = "";\n    for (const part of parts) {\n      current += "/" + part;\n      try {\n        if (!x2t.FS.analyzePath(current).exists) {\n          x2t.FS.mkdir(current);\n        }\n      } catch (err) {\n        console.error(err);\n      }\n    }\n  }\n  function writeFileMap(files) {\n    for (const [key, value] of Object.entries(files)) {\n      try {\n        const path = "/working/" + key.replace(/^\\/+/, "");\n        ensureDirectory(path.split("/").slice(0, -1).join("/"));\n        x2t.FS.writeFile(path, value);\n      } catch (err) {\n        console.error(key, err);\n      }\n    }\n  }\n  function writeFonts(fonts) {\n    cleanFonts();\n    if (!fonts) {\n      return;\n    }\n    for (const [key, value] of Object.entries(fonts)) {\n      try {\n        x2t.FS.writeFile("/working/fonts/" + key, value);\n      } catch (err) {\n        console.error(key, err);\n      }\n    }\n  }\n  function writePdfBin(pdfBin) {\n    const pdfPath = "/working/pdf.bin";\n    try {\n      if (x2t.FS.analyzePath(pdfPath).exists) {\n        x2t.FS.unlink(pdfPath);\n      }\n    } catch (err) {\n      console.error(err);\n    }\n    if (pdfBin?.byteLength) {\n      x2t.FS.writeFile(pdfPath, pdfBin);\n    }\n  }\n  function readMedia() {\n    const media = {};\n    try {\n      const files = x2t.FS.readdir("/working/media/");\n      for (const file of files) {\n        if (file !== "." && file !== "..") {\n          const fileData = x2t.FS.readFile("/working/media/" + file, {\n            encoding: "binary"\n          });\n          media[file] = fileData;\n        }\n      }\n    } catch (e) {\n      console.error(e);\n    }\n    return media;\n  }\n  function readDirectoryFiles(root, prefix) {\n    const files = {};\n    const visit = (dir, relDir) => {\n      try {\n        for (const file of x2t.FS.readdir(dir)) {\n          if (file === "." || file === "..") {\n            continue;\n          }\n          const path = `${dir}/${file}`;\n          const rel = relDir ? `${relDir}/${file}` : file;\n          const stat = x2t.FS.stat(path);\n          if (x2t.FS.isDir(stat.mode)) {\n            visit(path, rel);\n            continue;\n          }\n          files[`${prefix}/${rel}`] = x2t.FS.readFile(path, {\n            encoding: "binary"\n          });\n        }\n      } catch (err) {\n        console.error(err);\n      }\n    };\n    visit(root, "");\n    return files;\n  }\n  var xmlPath = "/working/params.xml";\n  function writeInputs({\n    fileFrom,\n    fileTo,\n    formatFrom,\n    formatTo,\n    data,\n    media,\n    pdfBin,\n    themes,\n    csvEncoding,\n    csvDelimiter,\n    csvDelimiterChar\n  }) {\n    const isCsvSource = formatFrom === 260 /* AVS_FILE_SPREADSHEET_CSV */ || fileFrom.toLowerCase().endsWith(".csv");\n    const usePdfBinPath = Boolean(pdfBin?.byteLength);\n    const params = {\n      m_sFileFrom: fileFrom,\n      m_sThemeDir: "/working/themes",\n      m_sFileTo: fileTo,\n      m_bIsPDFA: formatTo === 521 /* AVS_FILE_CROSSPLATFORM_PDFA */,\n      m_bIsNoBase64: usePdfBinPath ? false : true,\n      m_sFontDir: "/working/fonts/"\n    };\n    if (!usePdfBinPath) {\n      params.m_nFormatFrom = formatFrom ?? 0;\n      params.m_nFormatTo = formatTo ?? 0;\n    }\n    if (isCsvSource) {\n      params.m_nCsvTxtEncoding = csvEncoding ?? 46;\n      params.m_nCsvDelimiter = csvDelimiter ?? 4;\n      if (csvDelimiterChar) {\n        params.m_nCsvDelimiterChar = csvDelimiterChar;\n      }\n    }\n    const content = Object.entries(params).filter(([, v]) => v !== void 0 && v !== null && v !== "").reduce((a, [k, v]) => a + `<${k}>${v}</${k}>\n`, "");\n    const xml = `<?xml version="1.0" encoding="utf-8"?>\n<TaskQueueDataConvert\n  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\n  xmlns:xsd="http://www.w3.org/2001/XMLSchema"\n>\n${content}\n</TaskQueueDataConvert>`;\n    x2t.FS.writeFile(xmlPath, xml);\n    if (data) {\n      x2t.FS.writeFile(fileFrom, new Uint8Array(data));\n    }\n    if (media) {\n      cleanMedia();\n      for (const [key, value] of Object.entries(media)) {\n        try {\n          const data2 = /\\.(emf|svg|webp)$/i.test(key) ? transparentPng : value;\n          x2t.FS.writeFile("/working/" + key, data2);\n        } catch (err) {\n          console.error(key, err);\n        }\n      }\n    }\n    if (themes) {\n      cleanThemes();\n      writeFileMap(themes);\n    }\n  }\n  async function convert({\n    data,\n    fileFrom,\n    fileTo,\n    formatFrom,\n    formatTo,\n    media,\n    pdfBin,\n    fonts,\n    fontAliases,\n    fontExportAliases,\n    themes,\n    csvEncoding,\n    csvDelimiter,\n    csvDelimiterChar,\n    staticResource\n  }) {\n    let preparedData = data;\n    if (preparedData) {\n      preparedData = await sanitizeDocxForX2t(preparedData, fileFrom);\n      preparedData = await rewriteZipFontNames(preparedData, fileFrom, fontAliases);\n    }\n    const fromPath = "/working/" + fileFrom;\n    const toPath = "/working/" + fileTo;\n    const files = [fromPath, toPath, xmlPath];\n    if (pdfBin?.byteLength) {\n      files.push("/working/pdf.bin");\n    }\n    const needsPdfFonts = Boolean(pdfBin?.byteLength) || formatTo === 513 /* AVS_FILE_CROSSPLATFORM_PDF */ || formatTo === 521 /* AVS_FILE_CROSSPLATFORM_PDFA */;\n    let allFonts = fonts;\n    if (needsPdfFonts) {\n      const pdfFontsRoot = getWorkerStaticResource({\n        staticResource\n      }).x2t.pdfFonts.root;\n      const pdfFonts = await loadX2tPdfFonts(\n        self.location.origin,\n        pdfFontsRoot\n      );\n      allFonts = { ...pdfFonts, ...fonts };\n    }\n    writeFonts(allFonts);\n    writePdfBin(pdfBin);\n    writeInputs({\n      fileFrom: fromPath,\n      fileTo: toPath,\n      formatFrom,\n      formatTo,\n      data: preparedData,\n      media,\n      themes,\n      pdfBin,\n      csvEncoding,\n      csvDelimiter,\n      csvDelimiterChar\n    });\n    if (fileFrom.endsWith(".doc") || formatFrom == 66 /* AVS_FILE_DOCUMENT_DOC */) {\n      const viaPath = fromPath + ".docx";\n      try {\n        const pathInfo = x2t.FS.analyzePath(viaPath);\n        if (pathInfo.exists) {\n          x2t.FS.unlink(viaPath);\n        }\n      } catch (err) {\n      }\n      writeInputs({\n        fileFrom: fromPath,\n        fileTo: viaPath,\n        formatFrom: 66 /* AVS_FILE_DOCUMENT_DOC */,\n        formatTo: 65 /* AVS_FILE_DOCUMENT_DOCX */,\n        data: null\n      });\n      x2t.ccall("main1", ["number"], ["string"], [xmlPath]);\n      writeInputs({\n        fileFrom: viaPath,\n        fileTo: toPath,\n        formatFrom: 65 /* AVS_FILE_DOCUMENT_DOCX */,\n        formatTo,\n        data: null\n      });\n      files.push(viaPath);\n    }\n    try {\n      const pathInfo = x2t.FS.analyzePath(toPath);\n      if (pathInfo.exists) {\n        x2t.FS.unlink(toPath);\n      }\n    } catch (err) {\n    }\n    try {\n      x2t.ccall("main1", ["number"], ["string"], [xmlPath]);\n    } catch (e) {\n      console.error("ccall", e);\n    }\n    let output = null;\n    let readOutputError;\n    try {\n      output = x2t.FS.readFile(toPath);\n      if (output && fileTo === "Editor.bin" && fontAliases) {\n        output = restoreEditorBinFontNames(output, fontAliases);\n      }\n      if (output && hasOfficeZipExtension(fileTo) && fontExportAliases) {\n        const restored = await rewriteZipFontNames(\n          output.slice(0).buffer,\n          fileTo,\n          fontExportAliases\n        );\n        output = new Uint8Array(restored);\n      }\n    } catch (e) {\n      console.error("[x2t.worker] read output failed:", e);\n      readOutputError = serializeWorkerError(e);\n    }\n    if (!output) {\n      throw new X2tWorkerError("x2t conversion produced no output", {\n        stage: "read-output",\n        fileFrom,\n        fileTo,\n        readOutputError\n      });\n    }\n    const outputMedia = readMedia();\n    const outputThemes = readDirectoryFiles("/working/themes", "themes");\n    setTimeout(() => {\n      cleanupFiles(files);\n    });\n    return { output, media: outputMedia, themes: outputThemes };\n  }\n  self.onmessage = async (event) => {\n    const { id, type, payload } = event.data;\n    try {\n      switch (type) {\n        case "convert": {\n          await ensureInit(payload);\n          const result = await convert(payload);\n          const transferables = [];\n          if (result.output) {\n            transferables.push(result.output.buffer);\n          }\n          Object.values(result.media).forEach(\n            (m) => transferables.push(m.buffer)\n          );\n          self.postMessage(\n            { id, type: "convert:done", payload: result },\n            { transfer: transferables }\n          );\n          break;\n        }\n        default:\n          self.postMessage({\n            id,\n            type: "error",\n            error: `Unknown message type: ${type}`\n          });\n      }\n    } catch (error) {\n      self.postMessage({\n        id,\n        type: "error",\n        error: error instanceof Error ? error.message : String(error),\n        errorDetails: error instanceof X2tWorkerError ? error.details : void 0\n      });\n    }\n  };\n  self.postMessage({ type: "ready" });\n})();\n';

// src/client/sdk.js
function consoleWarn(message) {
  try {
    console.warn(`[host-base] ${message}`);
  } catch {
  }
}
var STATUS_PATH = "dsh-onlyoffice/status";
var STATUS_TIMEOUT_MS = 2e4;
var DEFAULT_SDK_BASE = "/dsh-onlyoffice/sdk";
var SDK_MARKER = "/onlyoffice/";
var OfficeViewError = class extends Error {
  constructor(stage, message, detail) {
    super(message);
    this.name = "OfficeViewError";
    this.stage = stage;
    this.detail = detail;
  }
};
var sdkPromise = null;
var sdkConfig = null;
var hostBase = null;
function hostBaseCandidates() {
  const out = [];
  const push = (value) => {
    if (typeof value !== "string") return;
    const clean = value.replace(/\/+$/, "");
    if (!out.includes(clean)) out.push(clean);
  };
  const origin = typeof location !== "undefined" ? location.origin : void 0;
  const base = typeof document !== "undefined" ? document.baseURI : void 0;
  try {
    if (typeof base === "string" && base !== "") {
      const dir = new URL(".", base);
      dir.search = "";
      dir.hash = "";
      push(dir.href);
    }
  } catch {
  }
  if (typeof origin === "string" && origin !== "" && origin !== "null" && origin !== "undefined") {
    push(origin);
  }
  push("");
  return out;
}
function setHostBase(base) {
  hostBase = typeof base === "string" && base !== "" ? base.replace(/\/+$/, "") : null;
}
async function probeHostBases() {
  const suffix = "/" + String(STATUS_PATH).replace(/^\/+/, "");
  const apiSuffix = "/dsh-onlyoffice/sdk/onlyoffice/9.4.0-develop/web-apps/apps/api/documents/api.js";
  const results = [];
  for (const base of hostBaseCandidates()) {
    const attempt = { base, status: null, api: null };
    for (const [key, path] of [["status", suffix], ["api", apiSuffix]]) {
      const url = base === "" ? path : `${base}${path}`;
      const started = typeof performance !== "undefined" ? performance.now() : Date.now();
      try {
        const response = await fetch(url, { credentials: "same-origin" });
        const body = await response.clone().text();
        attempt[key] = {
          url,
          status: response.status,
          contentType: response.headers.get("content-type"),
          ms: Math.round((typeof performance !== "undefined" ? performance.now() : Date.now()) - started),
          head: body.slice(0, 120).replace(/\s+/g, " ")
        };
      } catch (error) {
        attempt[key] = {
          url,
          error: error && error.name ? `${error.name}: ${error.message}` : String(error),
          ms: Math.round((typeof performance !== "undefined" ? performance.now() : Date.now()) - started)
        };
      }
    }
    results.push(attempt);
  }
  let baseURI = null;
  let origin = null;
  try {
    baseURI = typeof document !== "undefined" ? document.baseURI : null;
    origin = typeof location !== "undefined" ? location.origin : null;
  } catch {
  }
  return { baseURI, origin, hostBase, candidates: results };
}
function resolveHostUrl(ctx, path) {
  try {
    const service = ctx && typeof ctx.get === "function" ? ctx.get("betterSidebar") : void 0;
    if (service && typeof service.hostRouteUrl === "function") {
      return service.hostRouteUrl(path);
    }
  } catch {
  }
  const clean = String(path).replace(/^\/+/, "");
  const base = hostBase !== null ? hostBase : hostBaseCandidates()[0];
  return base === "" ? `/${clean}` : `${base}/${clean}`;
}
function timeoutSignal(timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  return {
    signal: controller.signal,
    dispose: () => clearTimeout(timer)
  };
}
async function fetchStatus(statusPath, url) {
  const timeout = timeoutSignal(STATUS_TIMEOUT_MS);
  let response;
  try {
    response = await fetch(url, {
      method: "GET",
      credentials: "same-origin",
      headers: { accept: "application/json" },
      signal: timeout.signal
    });
  } catch (error) {
    const reason = error && error.name === "AbortError" ? "timeout" : error?.message || String(error);
    throw new OfficeViewError("status", `GET ${statusPath} failed`, reason);
  } finally {
    timeout.dispose();
  }
  if (!response.ok) {
    throw new OfficeViewError("status", `GET ${statusPath} returned ${response.status}`, `HTTP ${response.status}`);
  }
  let status;
  try {
    status = await response.json();
  } catch (error) {
    throw new OfficeViewError("status", "Status response is not JSON", error?.message || String(error));
  }
  return status;
}
async function readStatus(ctx) {
  let explicit = null;
  try {
    const service = ctx && typeof ctx.get === "function" ? ctx.get("betterSidebar") : void 0;
    if (service && typeof service.hostRouteUrl === "function") {
      explicit = service.hostRouteUrl(STATUS_PATH);
    }
  } catch {
  }
  const suffix = "/" + String(STATUS_PATH).replace(/^\/+/, "");
  const urls = explicit ? [explicit] : hostBaseCandidates().map((base) => base === "" ? suffix : `${base}${suffix}`);
  const failures = [];
  for (const url of urls) {
    try {
      const status = await fetchStatus(STATUS_PATH, url);
      const base = url.slice(0, url.length - suffix.length);
      setHostBase(base);
      consoleWarn(`chosen base: ${base === "" ? "(root-relative)" : base} (from ${url})`);
      return { status, url };
    } catch (error) {
      failures.push({ url, error });
      consoleWarn(`candidate failed: ${url} -> ${error?.message || error}`);
    }
  }
  const first = failures[0];
  if (urls.length > 1) {
    throw new OfficeViewError(
      "status",
      `host route unreachable (${urls.length} bases tried)`,
      failures.map((failure) => `${failure.url}: ${failure.error?.message || failure.error}`).join(" | ")
    );
  }
  throw first ? first.error : new OfficeViewError("status", "no host route candidate", urls.join(", "));
}
function sdkOriginFrom(status) {
  const raw = status && typeof status.sdkBaseUrl === "string" && status.sdkBaseUrl.trim() ? status.sdkBaseUrl.trim() : DEFAULT_SDK_BASE;
  const clean = raw.replace(/^\/+/, "");
  const base = hostBase !== null ? hostBase : "";
  return base === "" ? `/${clean}` : `${base}/${clean}`;
}
function ensureSdkReady(ctx) {
  if (sdkConfig) return Promise.resolve(sdkConfig);
  if (sdkPromise) return sdkPromise;
  sdkPromise = (async () => {
    const { status } = await readStatus(ctx);
    if (status && status.ok === false && !status.version) {
      throw new OfficeViewError(
        "status",
        "The host status route reported a failure",
        typeof status.lastError === "string" && status.lastError ? status.lastError : "ok: false"
      );
    }
    const version = status && typeof status.version === "string" && status.version.trim() ? status.version.trim() : void 0;
    const cdnOrigin = sdkOriginFrom(status);
    registerOnlyOfficeStaticResource({
      cdnOrigin,
      ...version ? { onlyofficeVersion: version } : {}
    });
    const config = {
      status,
      version,
      cdnOrigin,
      sdkRoot: version ? `${cdnOrigin}${SDK_MARKER}${version}` : cdnOrigin
    };
    sdkConfig = config;
    return config;
  })();
  sdkPromise = sdkPromise.catch((error) => {
    sdkPromise = null;
    sdkConfig = null;
    throw error;
  });
  return sdkPromise;
}
async function saveDocumentFile(ctx, request) {
  const { sessionId, path, cwd, mtimeMs, bytes } = request || {};
  if (!sessionId || !path) {
    throw new OfficeViewError("save", "the document has no workspace path to save to", "missing sessionId/path");
  }
  const url = new URL(resolveHostUrl(ctx, "dsh-onlyoffice/save"));
  url.searchParams.set("sessionId", sessionId);
  url.searchParams.set("path", path);
  if (cwd) url.searchParams.set("cwd", cwd);
  if (Number.isFinite(mtimeMs)) url.searchParams.set("mtimeMs", String(mtimeMs));
  let response;
  try {
    response = await fetch(url.href, {
      method: "POST",
      credentials: "same-origin",
      headers: { "content-type": "application/octet-stream" },
      body: bytes
    });
  } catch (error) {
    throw new OfficeViewError("save", "the save request failed", error?.message || String(error));
  }
  let payload = null;
  try {
    payload = await response.json();
  } catch {
    payload = null;
  }
  if (!response.ok || !payload || payload.ok !== true) {
    const code = payload && payload.error ? payload.error.code : `http-${response.status}`;
    const message = payload && payload.error ? payload.error.message : `HTTP ${response.status}`;
    throw new OfficeViewError("save", message, `${code}`);
  }
  return payload;
}
function decodeVersionHeader(raw) {
  if (typeof raw !== "string" || raw === "") return void 0;
  try {
    const padded = raw.replace(/-/g, "+").replace(/_/g, "/");
    const binary = atob(padded + "=".repeat((4 - padded.length % 4) % 4));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return new TextDecoder().decode(bytes);
  } catch {
    return void 0;
  }
}
async function readDocumentFile(ctx, request) {
  const { sessionId, path, cwd, signal } = request || {};
  if (!sessionId || !path) {
    throw new OfficeViewError("read", "the document has no workspace path to read", "missing sessionId/path");
  }
  const url = new URL(resolveHostUrl(ctx, "dsh-onlyoffice/read"));
  url.searchParams.set("sessionId", sessionId);
  url.searchParams.set("path", path);
  if (cwd) url.searchParams.set("cwd", cwd);
  let response;
  try {
    response = await fetch(url.href, { credentials: "same-origin", signal });
  } catch (error) {
    if (error && error.name === "AbortError") throw error;
    throw new OfficeViewError("read", "the document read failed", error?.message || String(error));
  }
  if (!response.ok) {
    let detail = `HTTP ${response.status}`;
    try {
      const payload = await response.json();
      if (payload && payload.error) detail = `${payload.error.code}: ${payload.error.message}`;
    } catch {
    }
    throw new OfficeViewError("read", "the document could not be read", detail);
  }
  const bytes = await response.arrayBuffer();
  if (!bytes || bytes.byteLength === 0) {
    throw new OfficeViewError("read", "the document is empty", "empty body");
  }
  return {
    bytes,
    version: decodeVersionHeader(response.headers.get("x-dsh-file-version"))
  };
}
var workerFactoryInstalled = false;
function buildX2tWorker() {
  if (typeof Worker === "undefined" || typeof Blob === "undefined" || typeof URL === "undefined") {
    throw new OfficeViewError("engine", "Web Workers are not available", "Worker/Blob/URL missing");
  }
  const blob = new Blob([X2T_WORKER_SOURCE], { type: "text/javascript" });
  const blobUrl = URL.createObjectURL(blob);
  const worker = new Worker(blobUrl);
  let revoked = false;
  const revoke = () => {
    if (revoked) return;
    revoked = true;
    clearTimeout(timer);
    try {
      URL.revokeObjectURL(blobUrl);
    } catch {
    }
  };
  const timer = setTimeout(revoke, 3e4);
  worker.addEventListener("message", revoke, { once: true });
  worker.addEventListener("error", revoke, { once: true });
  return worker;
}
function installX2tWorkerFactory() {
  if (workerFactoryInstalled) return () => {
  };
  workerFactoryInstalled = true;
  setX2tWorkerFactory(buildX2tWorker);
  return () => {
    workerFactoryInstalled = false;
    setX2tWorkerFactory(null);
  };
}
var conversionChain = Promise.resolve();
var CONVERSION_WAIT_TIMEOUT_MS = 24e4;
function withConversionLock(task, waitTimeoutMs = CONVERSION_WAIT_TIMEOUT_MS) {
  let timer = null;
  const head = conversionChain.catch(() => void 0);
  const ready = head.then(() => "ready");
  const stalled = new Promise((resolve) => {
    timer = setTimeout(() => resolve("stalled"), waitTimeoutMs);
  });
  const run = Promise.race([ready, stalled]).then((state) => {
    if (state === "stalled") {
      console.warn(
        "[dsh-onlyoffice] the document conversion queue did not drain in time; running this conversion without the serialization guarantee"
      );
    }
    return task();
  });
  run.then(
    () => clearTimeout(timer),
    () => clearTimeout(timer)
  );
  conversionChain = run.then(
    () => void 0,
    () => void 0
  );
  return run;
}
function errorText(error) {
  if (!error) return "Unknown error";
  if (typeof error === "string") return error;
  const message = typeof error.message === "string" && error.message ? error.message : String(error);
  return message;
}
try {
  if (typeof window !== "undefined") window.__dshooProbe = probeHostBases;
} catch {
}

// src/client/theme.js
var DARK = "dark";
var LIGHT = "light";
var THEME_ATTRIBUTES = ["data-theme", "data-color-scheme", "data-dsw-theme", "data-mode"];
function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}
function srgbToLuminance(r, g, b) {
  const channel = (value) => {
    const v = clamp(value, 0, 255) / 255;
    return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}
function parseHexColor(text) {
  const match = /^#([0-9a-f]{3,8})$/i.exec(text);
  if (!match) return null;
  let hex = match[1];
  if (hex.length === 3 || hex.length === 4) {
    hex = hex.split("").map((ch) => ch + ch).join("");
  }
  if (hex.length !== 6 && hex.length !== 8) return null;
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16)
  };
}
function parseRgbColor(text) {
  const match = /^rgba?\(([^)]+)\)$/i.exec(text);
  if (!match) return null;
  const parts = match[1].replace(/\//g, " ").split(/[\s,]+/).filter(Boolean);
  if (parts.length < 3) return null;
  const channel = (raw) => {
    if (raw.endsWith("%")) return parseFloat(raw) / 100 * 255;
    return parseFloat(raw);
  };
  const r = channel(parts[0]);
  const g = channel(parts[1]);
  const b = channel(parts[2]);
  if (![r, g, b].every((value) => Number.isFinite(value))) return null;
  return { r, g, b };
}
function parseHslColor(text) {
  const match = /^hsla?\(([^)]+)\)$/i.exec(text);
  if (!match) return null;
  const parts = match[1].replace(/\//g, " ").replace(/,/g, " ").split(/\s+/).filter(Boolean);
  if (parts.length < 3) return null;
  const h = parseFloat(parts[0]);
  const s = parseFloat(parts[1]) / 100;
  const l = parseFloat(parts[2]) / 100;
  if (![h, s, l].every((value) => Number.isFinite(value))) return null;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const hp = (h % 360 + 360) % 360 / 60;
  const x = c * (1 - Math.abs(hp % 2 - 1));
  const [r1, g1, b1] = hp < 1 ? [c, x, 0] : hp < 2 ? [x, c, 0] : hp < 3 ? [0, c, x] : hp < 4 ? [0, x, c] : hp < 5 ? [x, 0, c] : [c, 0, x];
  const m = l - c / 2;
  return { r: (r1 + m) * 255, g: (g1 + m) * 255, b: (b1 + m) * 255 };
}
function parsePerceptualLightness(text) {
  const match = /^(?:oklch|oklab|lab|lch)\(\s*([\d.]+)(%?)/i.exec(text);
  if (!match) return null;
  let value = parseFloat(match[1]);
  if (!Number.isFinite(value)) return null;
  if (match[2] === "%") value /= 100;
  if (value > 1) value /= 100;
  return value;
}
function colorIsDark(raw) {
  const text = String(raw || "").trim().toLowerCase();
  if (!text) return null;
  if (text === "white" || text === "transparent") return false;
  if (text === "black") return true;
  const rgb = parseHexColor(text) || parseRgbColor(text) || parseHslColor(text);
  if (rgb) return srgbToLuminance(rgb.r, rgb.g, rgb.b) < 0.4;
  const lightness = parsePerceptualLightness(text);
  if (lightness !== null) return lightness < 0.55;
  return null;
}
function readTokenStyle(win, doc, token) {
  const roots = [doc.documentElement, doc.body].filter(Boolean);
  for (const root of roots) {
    try {
      const style = win.getComputedStyle(root);
      const value = style.getPropertyValue(token);
      if (value && value.trim()) return value.trim();
    } catch {
    }
  }
  return "";
}
function readColorScheme(win, doc) {
  const roots = [doc.documentElement, doc.body].filter(Boolean);
  for (const root of roots) {
    try {
      const scheme = String(win.getComputedStyle(root).colorScheme || "").toLowerCase();
      if (scheme.includes("dark")) return DARK;
      if (scheme.includes("light")) return LIGHT;
    } catch {
    }
  }
  return "";
}
function readAttributeSignal(doc) {
  const elements = [doc.documentElement, doc.body].filter(Boolean);
  for (const element of elements) {
    for (const name of THEME_ATTRIBUTES) {
      let value;
      try {
        value = element.getAttribute(name);
      } catch {
        continue;
      }
      if (!value) continue;
      const normalized = String(value).toLowerCase();
      if (normalized === DARK || normalized.endsWith("-dark")) return DARK;
      if (normalized === LIGHT || normalized.endsWith("-light")) return LIGHT;
    }
    try {
      const classes = element.classList;
      if (classes && typeof classes.contains === "function") {
        if (classes.contains("dark") || classes.contains("theme-dark")) return DARK;
        if (classes.contains("light") || classes.contains("theme-light")) return LIGHT;
      }
    } catch {
    }
  }
  return "";
}
function readMediaSignal(win) {
  try {
    const query = win.matchMedia?.("(prefers-color-scheme: dark)");
    if (query && typeof query.matches === "boolean") return query.matches ? DARK : LIGHT;
  } catch {
  }
  return "";
}
function isHostDark(win = typeof window === "undefined" ? void 0 : window) {
  if (!win || !win.document) return false;
  const doc = win.document;
  const tokenDark = colorIsDark(readTokenStyle(win, doc, "--dsw-alias-bg-base"));
  if (tokenDark !== null) return tokenDark;
  const labelDark = colorIsDark(readTokenStyle(win, doc, "--dsw-alias-label-primary"));
  if (labelDark !== null) return !labelDark;
  return readColorScheme(win, doc) === DARK || readAttributeSignal(doc) === DARK || readMediaSignal(win) === DARK;
}
function subscribeHostTheme(listener, win = typeof window === "undefined" ? void 0 : window) {
  if (!win || !win.document) return () => {
  };
  const doc = win.document;
  let current = isHostDark(win);
  let scheduled = false;
  const notify = () => {
    scheduled = false;
    const next = isHostDark(win);
    if (next === current) return;
    current = next;
    try {
      listener(next);
    } catch {
    }
  };
  const schedule = () => {
    if (scheduled) return;
    scheduled = true;
    try {
      win.requestAnimationFrame(notify);
    } catch {
      win.setTimeout(notify, 16);
    }
  };
  let observer;
  try {
    observer = new win.MutationObserver(schedule);
    const targets = [doc.documentElement, doc.body].filter(Boolean);
    for (const target of targets) {
      observer.observe(target, {
        attributes: true,
        attributeFilter: ["class", "style", ...THEME_ATTRIBUTES]
      });
    }
  } catch {
    observer = void 0;
  }
  let mediaQuery;
  let mediaListener;
  try {
    mediaQuery = win.matchMedia?.("(prefers-color-scheme: dark)");
    if (mediaQuery && typeof mediaQuery.addEventListener === "function") {
      mediaListener = () => schedule();
      mediaQuery.addEventListener("change", mediaListener);
    }
  } catch {
    mediaQuery = void 0;
  }
  return () => {
    try {
      observer?.disconnect();
    } catch {
    }
    try {
      if (mediaQuery && mediaListener) mediaQuery.removeEventListener("change", mediaListener);
    } catch {
    }
  };
}

// src/client/i18n.js
var DICT = {
  zh: {
    viewerOnlyOffice: "OnlyOffice 文档",
    viewerDocx: "Word 文档",
    viewerXlsx: "Excel 表格",
    viewerPptx: "PowerPoint 演示文稿",
    loadingEngine: "正在加载 OnlyOffice 引擎…",
    openingDocument: "正在打开文档…",
    converting: "正在转换文档…",
    slowHint: "首次打开需要下载并初始化 OnlyOffice 引擎，可能需要一点时间。",
    retry: "重试",
    reload: "重新加载",
    save: "保存",
    saving: "保存中…",
    savedNotice: "已保存到原文件。",
    saveFailed: "保存失败",
    saveConflict: "文件已在磁盘上被修改，请先「重新加载」再保存",
    saveHint: "把编辑后的文档写回原文件",
    saveReadOnly: "当前是只读模式，切换到编辑后再保存",
    externalChanged: "文件已在磁盘上更新。",
    externalReload: "重新加载",
    reloadFailed: "重新加载失败",
    saveNoTarget: "没有可写回的原始文件路径（仅支持从文件树打开的文档）",
    saveUnsupported: "该格式不支持直接写回（仅 docx / xlsx / pptx），请用「导出下载」",
    export: "导出下载",
    exporting: "导出中…",
    readOnly: "只读",
    editable: "编辑",
    toggleReadOnly: "切换只读 / 编辑",
    openExternal: "在浏览器中打开",
    dismiss: "关闭",
    exportingHint: "正在导出，请稍候…",
    loadFailed: "无法打开该文档",
    errorStatus: "无法连接 Office 预览服务（/dsh-onlyoffice/status）。",
    errorStatusHint: "宿主端 dsh-onlyoffice 插件可能未启用，或该路由暂时不可用。",
    errorNoMedia: "没有可用的文件地址（mediaUrl 缺失）。",
    errorUnsupported: "当前环境不支持 File API，无法打开 Office 文档。",
    errorFetch: "无法读取文件内容。",
    errorEngine: "OnlyOffice 引擎初始化失败。",
    errorConvert: "文档转换失败。",
    errorReady: "编辑器渲染超时。",
    errorTimeout: "加载超时。",
    errorTimeoutHint: "宿主端 SDK 路由可能不可达；可重试，或确认 dsh-onlyoffice 插件已启用。",
    errorRetryHint: "可以重试；若持续失败，请在文件树中下载该文件后用本地 Office 打开。",
    exportFailed: "导出失败",
    readonlyFailed: "切换只读状态失败"
  },
  en: {
    viewerOnlyOffice: "OnlyOffice document",
    viewerDocx: "Word document",
    viewerXlsx: "Excel spreadsheet",
    viewerPptx: "PowerPoint presentation",
    loadingEngine: "Loading the OnlyOffice engine…",
    openingDocument: "Opening the document…",
    converting: "Converting the document…",
    slowHint: "The first open downloads and initializes the OnlyOffice engine, so it may take a while.",
    retry: "Retry",
    reload: "Reload",
    save: "Save",
    saving: "Saving…",
    savedNotice: "Saved back to the original file.",
    saveFailed: "Save failed",
    saveConflict: "The file changed on disk — reload the document before saving",
    saveHint: "Write the edited document back to its file",
    saveReadOnly: "The document is read-only; switch to editing before saving",
    externalChanged: "The file changed on disk.",
    externalReload: "Reload",
    reloadFailed: "Reload failed",
    saveNoTarget: "No original file path to save to (open the document from the file tree)",
    saveUnsupported: "This format cannot be written back (docx / xlsx / pptx only); use Export instead",
    export: "Export",
    exporting: "Exporting…",
    readOnly: "Read-only",
    editable: "Editing",
    toggleReadOnly: "Toggle read-only / editing",
    openExternal: "Open in browser",
    dismiss: "Dismiss",
    exportingHint: "Exporting, please wait…",
    loadFailed: "Could not open this document",
    errorStatus: "Cannot reach the Office preview service (/dsh-onlyoffice/status).",
    errorStatusHint: "The host-side dsh-onlyoffice plugin may be disabled, or the route is temporarily unavailable.",
    errorNoMedia: "No file URL is available (mediaUrl is missing).",
    errorUnsupported: "This environment has no File API, so Office documents cannot be opened.",
    errorFetch: "Could not read the file contents.",
    errorEngine: "The OnlyOffice engine failed to initialize.",
    errorConvert: "The document conversion failed.",
    errorReady: "The editor did not finish rendering in time.",
    errorTimeout: "Loading timed out.",
    errorTimeoutHint: "The host SDK route may be unreachable; retry, or check that the dsh-onlyoffice plugin is enabled.",
    errorRetryHint: "You can retry; if it keeps failing, download the file from the tree and open it locally.",
    exportFailed: "Export failed",
    readonlyFailed: "Could not switch the read-only state"
  }
};
var FALLBACK_LANG = "en";
var activeLang = detectLang();
var listeners = /* @__PURE__ */ new Set();
function detectLang() {
  try {
    const nav = typeof navigator === "undefined" ? void 0 : navigator;
    const raw = String(nav && (nav.language || nav.languages?.[0]) || "");
    if (/^zh\b/i.test(raw)) return "zh";
    if (raw) return "en";
  } catch {
  }
  return FALLBACK_LANG;
}
function readLocaleLang(locale) {
  try {
    const snapshot = locale && typeof locale.getSnapshot === "function" ? locale.getSnapshot() : void 0;
    const id = snapshot && typeof snapshot.active === "string" ? snapshot.active : "";
    if (/^zh/i.test(id)) return "zh";
    if (/^en/i.test(id)) return "en";
    if (/^ja/i.test(id)) return "zh";
  } catch {
  }
  return void 0;
}
function lookup(lang, key) {
  const table = DICT[lang] || DICT[FALLBACK_LANG];
  return table[key] ?? DICT[FALLBACK_LANG][key] ?? key;
}
function t(key, params) {
  let value = lookup(activeLang, key);
  if (params) {
    for (const name of Object.keys(params)) {
      value = value.split(`{${name}}`).join(String(params[name]));
    }
  }
  return value;
}
function setLang(next) {
  if (!next || next === activeLang) return;
  activeLang = next;
  for (const listener of [...listeners]) {
    try {
      listener(activeLang);
    } catch {
    }
  }
}
function attachLocale(ctx) {
  const locale = ctx && ctx.locale;
  if (!locale || typeof locale.subscribe !== "function") return () => {
  };
  const sync = () => {
    const lang = readLocaleLang(locale);
    if (lang) setLang(lang);
  };
  sync();
  let unsubscribe;
  try {
    unsubscribe = locale.subscribe(sync);
  } catch {
    return () => {
    };
  }
  return () => {
    try {
      unsubscribe?.();
    } catch {
    }
  };
}

// src/client/diagnostics.js
var MAX_ENTRIES = 16;
var MAX_POSTS = 24;
var TEXT_LIMIT = 900;
var entries = [];
var installed = false;
var posts = 0;
var poster = null;
var reportChain = Promise.resolve();
function describeErrorEvent(event) {
  if (!event) return "unknown error";
  const parts = [];
  if (event.message) parts.push(String(event.message));
  if (event.filename) parts.push(`@ ${event.filename}${event.lineno ? `:${event.lineno}` : ""}${event.colno ? `:${event.colno}` : ""}`);
  if (event.error && event.error.stack) parts.push(String(event.error.stack));
  return parts.join(" ") || "unknown error";
}
function describeRejection(event) {
  const reason = event ? event.reason : void 0;
  if (!reason) return "unhandled rejection";
  if (typeof reason === "string") return reason;
  return reason.stack ? String(reason.stack) : String(reason.message || reason);
}
function describeResourceError(event) {
  const target = event ? event.target : void 0;
  if (!target || target === window) return "";
  const url = target.src || target.href || "";
  if (!url) return "";
  return `${target.tagName || "resource"} failed to load: ${url}`;
}
function recordDiagnostic(kind, text) {
  try {
    const entry = { at: Date.now(), kind: String(kind), text: String(text).slice(0, TEXT_LIMIT) };
    entries.push(entry);
    if (entries.length > MAX_ENTRIES) entries.shift();
    const payload = `[${entry.kind}] ${entry.text}`;
    if (poster && posts < MAX_POSTS) {
      posts += 1;
      reportChain = reportChain.then(() => poster(payload)).catch(() => {
      });
    }
  } catch {
  }
}
function diagnosticsDetail(extra) {
  const lines = [];
  if (extra) lines.push(String(extra));
  for (const entry of entries) lines.push(`[${entry.kind}] ${entry.text}`);
  return lines.join("\n");
}
function installDiagnostics(post) {
  if (typeof post === "function") poster = post;
  if (installed) return () => {
  };
  if (typeof window === "undefined" || typeof window.addEventListener !== "function") return () => {
  };
  installed = true;
  const onError = (event) => {
    const resource = describeResourceError(event);
    if (resource) {
      recordDiagnostic("resource", resource);
      return;
    }
    recordDiagnostic("page-error", describeErrorEvent(event));
  };
  const onRejection = (event) => recordDiagnostic("page-rejection", describeRejection(event));
  const consoleLikes = ["error", "warn"].map((level) => {
    const original = typeof console !== "undefined" ? console[level] : void 0;
    if (typeof original !== "function") return null;
    const wrapped = (...args) => {
      try {
        recordDiagnostic(
          `console-${level}`,
          args.map((value) => {
            if (value instanceof Error) return value.stack || value.message;
            if (typeof value === "string") return value;
            try {
              return JSON.stringify(value);
            } catch {
              return String(value);
            }
          }).join(" ").slice(0, TEXT_LIMIT)
        );
      } catch {
      }
      return original.apply(console, args);
    };
    console[level] = wrapped;
    return () => {
      try {
        console[level] = original;
      } catch {
      }
    };
  });
  const consoleRestore = [];
  for (const restore of consoleLikes) if (typeof restore === "function") consoleRestore.push(restore);
  try {
    window.addEventListener("error", onError, true);
    window.addEventListener("unhandledrejection", onRejection);
  } catch {
    installed = false;
    for (const restore of consoleRestore) restore();
    return () => {
    };
  }
  return () => {
    try {
      window.removeEventListener("error", onError, true);
      window.removeEventListener("unhandledrejection", onRejection);
    } catch {
    }
    for (const restore of consoleRestore) restore();
    installed = false;
  };
}
function environmentReport(containerId) {
  const lines = [];
  try {
    lines.push(`page: ${location.href} (origin=${location.origin})`);
  } catch {
    lines.push("page: unavailable");
  }
  try {
    const api = typeof window !== "undefined" ? window.DocsAPI : void 0;
    lines.push(`DocsAPI: ${api ? "present" : "absent"}`);
  } catch {
    lines.push("DocsAPI: unreadable");
  }
  try {
    const host = document.getElementById(containerId);
    const frames = host ? host.querySelectorAll("iframe") : [];
    const wrapper = document.querySelector(`[data-onlyoffice-container-id="${containerId}"]`);
    const wrapperFrames = wrapper ? wrapper.querySelectorAll("iframe") : [];
    lines.push(
      `container: ${host ? "present" : "MISSING"} (${frames.length} iframe(s)); wrapper iframes: ${wrapperFrames.length}${host ? "" : " (DocsAPI replaced the id element)"}`
    );
    for (const frame of Array.from(frames).slice(0, 4)) lines.push(`  frame: ${frame.src || "(no src)"}`);
  } catch (error) {
    lines.push(`container: probe failed (${error && error.message ? error.message : String(error)})`);
  }
  try {
    const preload = document.querySelector("iframe[data-onlyoffice-preload]");
    lines.push(`preload iframe: ${preload ? preload.src : "none"}`);
    lines.push(`document iframes: ${document.querySelectorAll("iframe").length}`);
  } catch {
  }
  return lines.join("\n");
}
function attachEditorDiagnostics(containerId, options = {}) {
  if (typeof document === "undefined") return () => {
  };
  let timer = null;
  let ticks = 0;
  let attachedWindow = null;
  const listeners2 = [];
  const keyHandler = typeof options.onKeyDown === "function" ? options.onKeyDown : null;
  const onIframeError = (event) => recordDiagnostic("iframe-error", describeErrorEvent(event));
  const onIframeRejection = (event) => recordDiagnostic("iframe-rejection", describeRejection(event));
  const onIframeResource = (event) => {
    const text = describeResourceError(event);
    if (text) recordDiagnostic("iframe-resource", text);
  };
  const tryAttach = () => {
    if (attachedWindow) return true;
    let frame = null;
    try {
      const host = document.getElementById(containerId);
      frame = host ? host.querySelector("iframe") : null;
    } catch {
      frame = null;
    }
    if (!frame) return false;
    let win = null;
    try {
      win = frame.contentWindow;
    } catch {
      win = null;
    }
    if (!win || typeof win.addEventListener !== "function") return false;
    try {
      win.addEventListener("error", onIframeError, true);
      win.addEventListener("unhandledrejection", onIframeRejection);
      win.addEventListener("error", onIframeResource, true);
      if (keyHandler) win.addEventListener("keydown", keyHandler, true);
      listeners2.push(() => {
        try {
          win.removeEventListener("error", onIframeError, true);
          win.removeEventListener("unhandledrejection", onIframeRejection);
          win.removeEventListener("error", onIframeResource, true);
          if (keyHandler) win.removeEventListener("keydown", keyHandler, true);
        } catch {
        }
      });
      attachedWindow = win;
      return true;
    } catch {
      return false;
    }
  };
  if (!tryAttach()) {
    timer = setInterval(() => {
      ticks += 1;
      if (tryAttach() || ticks > 480) {
        clearInterval(timer);
        timer = null;
      }
    }, 250);
    if (timer && typeof timer === "object" && typeof timer.unref === "function") timer.unref();
  }
  return () => {
    if (timer !== null) clearInterval(timer);
    timer = null;
    for (const remove of listeners2.splice(0)) remove();
    attachedWindow = null;
  };
}

// src/client/shortcut.js
function isSaveShortcut(event) {
  if (!event) return false;
  if (!(event.metaKey === true || event.ctrlKey === true)) return false;
  if (event.altKey === true || event.shiftKey === true) return false;
  return event.key === "s" || event.key === "S";
}
function saveShortcutLabel() {
  try {
    const platform = typeof navigator === "undefined" ? "" : String(navigator.platform || navigator.userAgent || "");
    if (/mac|iphone|ipad|ipod/i.test(platform)) return "⌘S";
  } catch {
  }
  return "Ctrl+S";
}

// src/client/bytes.js
function toUint8Array(value) {
  if (!value) return null;
  if (value instanceof Uint8Array) return value;
  if (typeof ArrayBuffer !== "undefined" && value instanceof ArrayBuffer) return new Uint8Array(value);
  if (typeof ArrayBuffer !== "undefined" && ArrayBuffer.isView && ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  return null;
}
function sameBytes(left, right) {
  const a = toUint8Array(left);
  const b = toUint8Array(right);
  if (!a || !b) return false;
  if (a.byteLength !== b.byteLength) return false;
  if (a.byteLength === 0) return true;
  for (let index = 0; index < a.byteLength; index += 1) {
    if (a[index] !== b[index]) return false;
  }
  return true;
}

// src/client/autosave.js
var AUTOSAVE_DELAY_MS = 1800;
function createAutoSaver(options) {
  const run = options && typeof options.run === "function" ? options.run : () => {
  };
  const delayMs = Number.isFinite(options && options.delayMs) ? Math.max(100, Math.trunc(options.delayMs)) : AUTOSAVE_DELAY_MS;
  const setTimer = options && options.setTimer || ((fn, ms) => setTimeout(fn, ms));
  const clearTimer = options && options.clearTimer || ((handle) => clearTimeout(handle));
  let timer = null;
  let running = false;
  let disposed = false;
  const cancel = () => {
    if (timer !== null) {
      clearTimer(timer);
      timer = null;
    }
  };
  const flush = async () => {
    cancel();
    if (disposed || running) return;
    running = true;
    try {
      await run();
    } finally {
      running = false;
    }
  };
  const schedule = () => {
    if (disposed) return;
    cancel();
    timer = setTimer(() => {
      timer = null;
      void flush();
    }, delayMs);
  };
  return {
    schedule,
    flush,
    cancel,
    dispose() {
      disposed = true;
      cancel();
    },
    isPending: () => timer !== null,
    isRunning: () => running
  };
}

// src/client/office-view.jsx
var import_jsx_runtime = require("react/jsx-runtime");
var DOCUMENT_READY_TIMEOUT_MS = 18e4;
var LOAD_WATCHDOG_MS = 24e4;
var SLOW_HINT_DELAY_MS = 12e3;
var NOTICE_TIMEOUT_MS = 8e3;
var KIND_TO_FILE_TYPE = {
  docx: FILE_TYPE.DOCX,
  xlsx: FILE_TYPE.XLSX,
  pptx: FILE_TYPE.PPTX
};
var KIND_TO_EXT = {
  docx: "docx",
  xlsx: "xlsx",
  pptx: "pptx"
};
var MIME_BY_EXT = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  doc: "application/msword",
  odt: "application/vnd.oasis.opendocument.text",
  rtf: "application/rtf",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xls: "application/vnd.ms-excel",
  ods: "application/vnd.oasis.opendocument.spreadsheet",
  csv: "text/csv",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ppt: "application/vnd.ms-powerpoint",
  odp: "application/vnd.oasis.opendocument.presentation"
};
var containerSeq = 0;
var savedBytesByTarget = /* @__PURE__ */ new Map();
var SAVED_BYTES_LIMIT = 8;
function savedBytesKey(sessionId, path) {
  return `${String(sessionId || "")}\0${String(path || "")}`;
}
function rememberSavedBytes(sessionId, path, bytes) {
  if (!sessionId || !path) return;
  const key = savedBytesKey(sessionId, path);
  savedBytesByTarget.delete(key);
  savedBytesByTarget.set(key, bytes);
  while (savedBytesByTarget.size > SAVED_BYTES_LIMIT) {
    const oldest = savedBytesByTarget.keys().next();
    if (oldest.done === true) break;
    savedBytesByTarget.delete(oldest.value);
  }
}
function knownSavedBytes(sessionId, path) {
  if (!sessionId || !path) return null;
  return savedBytesByTarget.get(savedBytesKey(sessionId, path)) ?? null;
}
var WATCHER_RELOAD_WINDOW_MS = 1e4;
var reloadExpectedUntil = /* @__PURE__ */ new Map();
function markWatcherReloadExpected(sessionId, path) {
  if (!sessionId || !path) return;
  reloadExpectedUntil.set(savedBytesKey(sessionId, path), Date.now() + WATCHER_RELOAD_WINDOW_MS);
  while (reloadExpectedUntil.size > SAVED_BYTES_LIMIT) {
    const oldest = reloadExpectedUntil.keys().next();
    if (oldest.done === true) break;
    reloadExpectedUntil.delete(oldest.value);
  }
}
function isWatcherReloadExpected(sessionId, path) {
  if (!sessionId || !path) return false;
  const until = reloadExpectedUntil.get(savedBytesKey(sessionId, path));
  if (until === void 0) return false;
  if (Date.now() > until) {
    reloadExpectedUntil.delete(savedBytesKey(sessionId, path));
    return false;
  }
  return true;
}
function clearWatcherReloadExpected(sessionId, path) {
  if (!sessionId || !path) return;
  reloadExpectedUntil.delete(savedBytesKey(sessionId, path));
}
function nextContainerId() {
  containerSeq += 1;
  const random = Math.random().toString(36).slice(2, 8);
  return `dsh-onlyoffice-editor-${containerSeq}-${random}`;
}
function pathBaseName(path) {
  const text = String(path || "");
  const parts = text.split(/[\\/]/);
  return parts[parts.length - 1] || "";
}
function fileExtension(name) {
  const match = /\.([A-Za-z0-9]{1,8})$/.exec(String(name || ""));
  return match ? match[1].toLowerCase() : "";
}
function resolveFileName(title, path, kind) {
  const fallbackExt = KIND_TO_EXT[kind] || "docx";
  const fromTitle = typeof title === "string" ? title.trim() : "";
  const fromPath = pathBaseName(path);
  const raw = fromTitle || fromPath || `document.${fallbackExt}`;
  return fileExtension(raw) ? raw : `${raw}.${fallbackExt}`;
}
function mimeFor(fileName) {
  return MIME_BY_EXT[fileExtension(fileName)] || "application/octet-stream";
}
function disposeManager(manager, containerId) {
  try {
    manager?.setSaveShortcutHandler?.(null);
    manager?.setChangeHandler?.(null);
  } catch {
  }
  try {
    manager?.destroy?.();
  } catch (error) {
    console.warn("[dsh-onlyoffice] failed to destroy the Office editor", error);
  }
  try {
    editorManagerFactory.destroy(containerId);
  } catch (error) {
    console.warn("[dsh-onlyoffice] failed to release the editor manager", error);
  }
}
function fileFromBytes(bytes, fileName) {
  if (typeof File !== "function") {
    throw new OfficeViewError("fetch", t("errorUnsupported"));
  }
  return new File([bytes], fileName, { type: mimeFor(fileName) });
}
async function fetchOfficeFile(mediaUrl, fileName, signal) {
  if (!mediaUrl) {
    throw new OfficeViewError("fetch", t("errorNoMedia"));
  }
  if (typeof File !== "function") {
    throw new OfficeViewError("fetch", t("errorUnsupported"));
  }
  let response;
  try {
    response = await fetch(mediaUrl, { credentials: "same-origin", signal });
  } catch (error) {
    if (error && error.name === "AbortError") throw error;
    throw new OfficeViewError("fetch", t("errorFetch"), errorText(error));
  }
  if (!response.ok) {
    throw new OfficeViewError("fetch", t("errorFetch"), `HTTP ${response.status} ${response.statusText || ""}`.trim());
  }
  const buffer = await response.arrayBuffer();
  if (!buffer || buffer.byteLength === 0) {
    throw new OfficeViewError("fetch", t("errorFetch"), "empty response body");
  }
  return new File([buffer], fileName, { type: mimeFor(fileName) });
}
function waitForDocumentReady(containerId, timeoutMs, signal) {
  return new Promise((resolve, reject) => {
    if (signal && signal.aborted) {
      reject(new OfficeViewError("aborted", "aborted"));
      return;
    }
    let settled = false;
    let timer = null;
    const cleanup = () => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      if (timer !== null) clearTimeout(timer);
      try {
        signal?.removeEventListener?.("abort", onAbort);
      } catch {
      }
    };
    const finish = (settle, value) => {
      if (settled) return;
      settled = true;
      cleanup();
      settle(value);
    };
    const handler = (data) => {
      if (data && data.instanceId && data.instanceId !== containerId) return;
      finish(resolve, data);
    };
    const onAbort = () => finish(reject, new OfficeViewError("aborted", "aborted"));
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
    timer = setTimeout(
      () => finish(reject, new OfficeViewError("ready", t("errorReady"))),
      timeoutMs
    );
    try {
      signal?.addEventListener?.("abort", onAbort, { once: true });
    } catch {
    }
  });
}
function editorFrameInfo(containerId) {
  return environmentReport(containerId);
}
function toViewerError(error) {
  if (error instanceof OfficeViewError) {
    return { stage: error.stage, message: error.message, detail: error.detail };
  }
  return { stage: "engine", message: t("errorEngine"), detail: errorText(error) };
}
function stageHint(stage) {
  if (stage === "status") return t("errorStatusHint");
  if (stage === "timeout") return t("errorTimeoutHint");
  return t("errorRetryHint");
}
function stageMessage(stage, fallback) {
  if (stage === "status") return t("errorStatus");
  return fallback;
}
function hasEditorFrame(containerId) {
  try {
    const host = document.getElementById(containerId);
    if (host && host.querySelector("iframe")) return true;
    const wrapper = document.querySelector(`[data-onlyoffice-container-id="${containerId}"]`);
    if (wrapper && wrapper.querySelector("iframe")) return true;
    return Array.from(document.querySelectorAll('iframe[name="frameEditor"]')).some(
      (frame) => String(frame.src || "").includes(containerId)
    );
  } catch {
    return false;
  }
}
function OfficeView(props) {
  const { path, title, mediaUrl, kind, bytes, waitingForBytes } = props;
  const ctxRef = (0, import_react.useRef)(props.ctx);
  ctxRef.current = props.ctx;
  const [containerId] = (0, import_react.useState)(() => nextContainerId());
  const [phase, setPhase] = (0, import_react.useState)("loading");
  const [error, setError] = (0, import_react.useState)(null);
  const [readOnly, setReadOnly] = (0, import_react.useState)(false);
  const [busy, setBusy] = (0, import_react.useState)("");
  const [notice, setNotice] = (0, import_react.useState)("");
  const [reloadToken, setReloadToken] = (0, import_react.useState)(0);
  const [slowHint, setSlowHint] = (0, import_react.useState)(false);
  const [dark, setDark] = (0, import_react.useState)(() => isHostDark());
  const [externalUpdate, setExternalUpdate] = (0, import_react.useState)(null);
  const externalUpdateRef = (0, import_react.useRef)(null);
  externalUpdateRef.current = externalUpdate;
  const managerRef = (0, import_react.useRef)(null);
  const mountedRef = (0, import_react.useRef)(true);
  const readOnlyRef = (0, import_react.useRef)(readOnly);
  const themeRef = (0, import_react.useRef)(dark);
  const phaseRef = (0, import_react.useRef)(phase);
  const noticeTimerRef = (0, import_react.useRef)(null);
  readOnlyRef.current = readOnly;
  themeRef.current = dark;
  phaseRef.current = phase;
  const fileName = (0, import_react.useMemo)(
    () => resolveFileName(title, path, kind),
    [title, path, kind]
  );
  const fileNameRef = (0, import_react.useRef)(fileName);
  fileNameRef.current = fileName;
  const ownerScope = props.scope || {};
  const saveSessionId = props.sessionId || ownerScope.sessionId || "";
  const savePath = props.sourcePath || props.path || "";
  const saveCwd = props.cwd || ownerScope.cwd || "";
  const nativeSaveable = /^(docx|xlsx|pptx)$/i.test((savePath.match(/\.([A-Za-z0-9]+)$/) || [])[1] || "");
  const hasSaveTarget = Boolean(saveSessionId) && Boolean(savePath);
  const canSaveTarget = hasSaveTarget && nativeSaveable;
  const initialMtime = Number.isFinite(props.mtimeMs) ? props.mtimeMs : void 0;
  const savedMtimeRef = (0, import_react.useRef)(initialMtime);
  const saveRef = (0, import_react.useRef)(() => {
  });
  const savingRef = (0, import_react.useRef)(false);
  const lastSavedBytesRef = (0, import_react.useRef)(null);
  const autoSaveRef = (0, import_react.useRef)(() => {
  });
  const autoSaverRef = (0, import_react.useRef)(null);
  if (autoSaverRef.current === null) {
    autoSaverRef.current = createAutoSaver({
      delayMs: AUTOSAVE_DELAY_MS,
      run: () => autoSaveRef.current()
    });
  }
  const ignoreChangesUntilRef = (0, import_react.useRef)(0);
  const lastRenderedVersionRef = (0, import_react.useRef)(void 0);
  const deferredExternalCheckRef = (0, import_react.useRef)(null);
  const forceRebuildRef = (0, import_react.useRef)(false);
  const rendererRef = (0, import_react.useRef)(props.renderer);
  rendererRef.current = props.renderer;
  const rendererRevision = props.renderer ? props.renderer.revision : void 0;
  const interceptSaveKey = (0, import_react.useCallback)((event) => {
    if (!isSaveShortcut(event)) return;
    event.preventDefault();
    event.stopImmediatePropagation?.();
    event.stopPropagation();
    saveRef.current();
  }, []);
  const clearNoticeTimer = (0, import_react.useCallback)(() => {
    if (noticeTimerRef.current !== null) {
      clearTimeout(noticeTimerRef.current);
      noticeTimerRef.current = null;
    }
  }, []);
  const showNotice = (0, import_react.useCallback)(
    (text) => {
      clearNoticeTimer();
      setNotice(text);
      noticeTimerRef.current = setTimeout(() => {
        noticeTimerRef.current = null;
        if (mountedRef.current) setNotice("");
      }, NOTICE_TIMEOUT_MS);
    },
    [clearNoticeTimer]
  );
  (0, import_react.useEffect)(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      clearNoticeTimer();
      autoSaverRef.current?.dispose();
      if (deferredExternalCheckRef.current !== null) clearTimeout(deferredExternalCheckRef.current);
      deferredExternalCheckRef.current = null;
    };
  }, [clearNoticeTimer]);
  (0, import_react.useEffect)(() => subscribeHostTheme((nextDark) => setDark(nextDark)), []);
  (0, import_react.useEffect)(() => {
    const onSaved = (data) => {
      if (data && data.instanceId && data.instanceId !== containerId) return;
      saveRef.current();
    };
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.ONSAVE, onSaved);
    return () => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.ONSAVE, onSaved);
    };
  }, [containerId]);
  (0, import_react.useEffect)(() => {
    const rendererMode = rendererRevision !== void 0;
    const incomingBytes = rendererMode ? null : toUint8Array(bytes);
    const knownSaved = lastSavedBytesRef.current ?? knownSavedBytes(saveSessionId, savePath);
    const sameAsSaved = incomingBytes !== null && knownSaved !== null && sameBytes(incomingBytes, knownSaved);
    const ownWatcherReload = managerRef.current !== null && !rendererMode && waitingForBytes === true && isWatcherReloadExpected(saveSessionId, savePath);
    const ownRendererReload = rendererMode && managerRef.current !== null && lastRenderedVersionRef.current !== void 0 && isWatcherReloadExpected(saveSessionId, savePath);
    if (ownRendererReload) {
      try {
        rendererRef.current?.loaded?.(lastRenderedVersionRef.current);
      } catch {
      }
      if (deferredExternalCheckRef.current === null) {
        deferredExternalCheckRef.current = setTimeout(async () => {
          deferredExternalCheckRef.current = null;
          try {
            const loaded = await rendererRef.current?.load?.();
            if (loaded && loaded.version && loaded.version !== lastRenderedVersionRef.current) {
              clearWatcherReloadExpected(saveSessionId, savePath);
              rendererRef.current?.reload?.();
            }
          } catch {
          }
        }, 1500);
      }
      const detach = attachEditorDiagnostics(containerId, { onKeyDown: interceptSaveKey });
      return () => detach();
    }
    if (managerRef.current !== null && (sameAsSaved || ownWatcherReload)) {
      if (sameAsSaved) {
        lastSavedBytesRef.current = incomingBytes;
        clearWatcherReloadExpected(saveSessionId, savePath);
      }
      const detach = attachEditorDiagnostics(containerId, { onKeyDown: interceptSaveKey });
      return () => detach();
    }
    clearWatcherReloadExpected(saveSessionId, savePath);
    let cancelled = false;
    const controller = new AbortController();
    if (rendererMode && !forceRebuildRef.current && managerRef.current !== null && lastRenderedVersionRef.current !== void 0 && !waitingForBytes) {
      const inPlaceController = new AbortController();
      let inPlaceCancelled = false;
      const detachInPlace = attachEditorDiagnostics(containerId, {
        onKeyDown: interceptSaveKey
      });
      const refreshInPlace = async () => {
        try {
          const loaded = await rendererRef.current.load(inPlaceController.signal);
          if (inPlaceCancelled) return;
          if (loaded.version !== void 0 && loaded.version === lastRenderedVersionRef.current) {
            rendererRef.current?.loaded?.(loaded.version);
            return;
          }
          if (loaded.version !== void 0) {
            try {
              rendererRef.current?.loaded?.(loaded.version);
            } catch {
            }
          }
          clearWatcherReloadExpected(saveSessionId, savePath);
          if (!inPlaceCancelled) {
            setExternalUpdate({ version: loaded.version, bytes: loaded.bytes });
          }
        } catch {
        }
      };
      void refreshInPlace();
      return () => {
        inPlaceCancelled = true;
        inPlaceController.abort();
        detachInPlace();
      };
    }
    forceRebuildRef.current = false;
    setExternalUpdate(null);
    autoSaverRef.current?.cancel();
    disposeManager(managerRef.current, containerId);
    managerRef.current = null;
    savedMtimeRef.current = initialMtime;
    setPhase("loading");
    setError(null);
    setNotice("");
    setBusy("");
    setSlowHint(false);
    const slowTimer = setTimeout(() => {
      if (!cancelled && mountedRef.current) setSlowHint(true);
    }, SLOW_HINT_DELAY_MS);
    const watchdogTimer = setTimeout(() => {
      if (cancelled || !mountedRef.current) return;
      const current = phaseRef.current;
      if (current === "ready" || current === "error") return;
      recordDiagnostic("viewer-timeout", editorFrameInfo(containerId));
      setError({
        stage: "timeout",
        message: t("errorTimeout"),
        detail: diagnosticsDetail(`${t("errorTimeoutHint")}
${editorFrameInfo(containerId)}`)
      });
      setPhase("error");
    }, LOAD_WATCHDOG_MS);
    const detachEditorDiagnostics = attachEditorDiagnostics(containerId, {
      onKeyDown: interceptSaveKey
    });
    const run = async () => {
      try {
        if (waitingForBytes) {
          setPhase("loading");
          return;
        }
        await ensureSdkReady(ctxRef.current);
        if (cancelled) return;
        setPhase("opening");
        let file;
        let loadedVersion;
        if (rendererMode) {
          const loaded = await rendererRef.current.load(controller.signal);
          if (cancelled) return;
          file = fileFromBytes(loaded.bytes, fileNameRef.current);
          loadedVersion = loaded.version;
        } else if (bytes) {
          file = fileFromBytes(bytes, fileNameRef.current);
        } else {
          file = await fetchOfficeFile(mediaUrl, fileNameRef.current, controller.signal);
        }
        if (cancelled) return;
        setPhase("converting");
        const manager = await withConversionLock(async () => {
          const created = await OnlyOfficeManager.createWithFile(
            {
              containerId,
              fileType: KIND_TO_FILE_TYPE[kind] || FILE_TYPE.DOCX,
              defaultFileName: fileNameRef.current,
              readOnly: readOnlyRef.current,
              theme: themeRef.current ? OFFICE_THEME.DARK : OFFICE_THEME.WHITE
            },
            file
          );
          if (cancelled) {
            disposeManager(created, containerId);
            return null;
          }
          return created;
        });
        if (!manager || cancelled) return;
        managerRef.current = manager;
        try {
          manager.setSaveShortcutHandler?.(() => saveRef.current());
          manager.setChangeHandler?.(() => {
            if (ignoreChangesUntilRef.current > Date.now()) return;
            autoSaverRef.current?.schedule();
          });
        } catch (error2) {
          console.warn("[dsh-onlyoffice] could not hook the editor save affordances", error2);
        }
        setReadOnly(Boolean(manager.getReadOnly?.()));
        try {
          await waitForDocumentReady(containerId, DOCUMENT_READY_TIMEOUT_MS, controller.signal);
        } catch (readyError) {
          if (cancelled || controller.signal.aborted) return;
          if (hasEditorFrame(containerId)) {
            showNotice(t("errorReady"));
          } else {
            throw readyError;
          }
        }
        if (cancelled) return;
        setPhase("ready");
        setSlowHint(false);
        if (rendererMode && loadedVersion !== void 0) {
          try {
            rendererRef.current?.loaded?.(loadedVersion);
          } catch (error2) {
            console.warn("[dsh-onlyoffice] could not report the rendered version", error2);
          }
          lastRenderedVersionRef.current = loadedVersion;
        }
      } catch (loadError) {
        if (cancelled || controller.signal.aborted) return;
        disposeManager(managerRef.current, containerId);
        managerRef.current = null;
        const viewerError = toViewerError(loadError);
        recordDiagnostic(
          "viewer-error",
          `${viewerError.stage}: ${viewerError.message}${viewerError.detail ? ` — ${viewerError.detail}` : ""} (${editorFrameInfo(containerId)})`
        );
        setError({ ...viewerError, detail: diagnosticsDetail(viewerError.detail) });
        setPhase("error");
      } finally {
        clearTimeout(slowTimer);
        clearTimeout(watchdogTimer);
      }
    };
    run();
    return () => {
      cancelled = true;
      controller.abort();
      clearTimeout(slowTimer);
      clearTimeout(watchdogTimer);
      detachEditorDiagnostics();
    };
  }, [containerId, kind, mediaUrl, bytes, waitingForBytes, reloadToken, showNotice, initialMtime, interceptSaveKey, rendererRevision]);
  (0, import_react.useEffect)(() => {
    return () => {
      disposeManager(managerRef.current, containerId);
      managerRef.current = null;
    };
  }, [containerId]);
  (0, import_react.useEffect)(() => {
    const manager = managerRef.current;
    if (!manager || typeof manager.setTheme !== "function") return;
    const next = dark ? OFFICE_THEME.DARK : OFFICE_THEME.WHITE;
    let current;
    try {
      current = manager.getTheme?.();
    } catch {
      current = void 0;
    }
    if (current === next) return;
    try {
      const pending = manager.setTheme(next);
      pending?.catch?.((themeError) => {
        console.warn("[dsh-onlyoffice] failed to sync the editor theme", themeError);
      });
    } catch (themeError) {
      console.warn("[dsh-onlyoffice] failed to sync the editor theme", themeError);
    }
  }, [dark]);
  const toggleReadOnly = (0, import_react.useCallback)(async () => {
    const manager = managerRef.current;
    if (!manager) return;
    setBusy("readonly");
    try {
      await manager.setReadOnly(!manager.getReadOnly());
      if (mountedRef.current) setReadOnly(Boolean(manager.getReadOnly()));
    } catch (toggleError) {
      if (mountedRef.current) showNotice(`${t("readonlyFailed")}: ${errorText(toggleError)}`);
    } finally {
      if (mountedRef.current) setBusy("");
    }
  }, [showNotice]);
  const exportDocument = (0, import_react.useCallback)(async () => {
    const manager = managerRef.current;
    if (!manager) return;
    setBusy("export");
    try {
      await withConversionLock(() => manager.downloadExport());
    } catch (exportError) {
      if (mountedRef.current) showNotice(`${t("exportFailed")}: ${errorText(exportError)}`);
    } finally {
      if (mountedRef.current) setBusy("");
    }
  }, [showNotice]);
  const saveToFile = (0, import_react.useCallback)(async (silent = false) => {
    const manager = managerRef.current;
    if (!manager || !canSaveTarget || savingRef.current) return;
    savingRef.current = true;
    setBusy("save");
    try {
      const exported = await withConversionLock(() => manager.exportAsBlob());
      if (!exported || !exported.blob) {
        throw new OfficeViewError("save", t("saveFailed"), "the engine produced no document bytes");
      }
      const bytes2 = await exported.blob.arrayBuffer();
      const result = await saveDocumentFile(ctxRef.current, {
        sessionId: saveSessionId,
        path: savePath,
        cwd: saveCwd || void 0,
        mtimeMs: savedMtimeRef.current,
        bytes: bytes2
      });
      const savedSnapshot = new Uint8Array(bytes2.slice(0));
      lastSavedBytesRef.current = savedSnapshot;
      rememberSavedBytes(saveSessionId, savePath, savedSnapshot);
      markWatcherReloadExpected(saveSessionId, savePath);
      if (result.version !== void 0 && result.version !== null && result.version !== "") {
        const version = String(result.version);
        lastRenderedVersionRef.current = version;
        try {
          rendererRef.current?.loaded?.(version);
        } catch (error2) {
          console.warn("[dsh-onlyoffice] could not report the saved version", error2);
        }
      }
      if (Number.isFinite(result.mtimeMs)) savedMtimeRef.current = result.mtimeMs;
      ignoreChangesUntilRef.current = Date.now() + 800;
      if (!silent && mountedRef.current) showNotice(t("savedNotice"));
    } catch (saveError) {
      const code = saveError && saveError.detail ? String(saveError.detail) : "";
      const isConflict = code.includes("conflict");
      if (mountedRef.current) {
        showNotice(`${isConflict ? t("saveConflict") : t("saveFailed")}: ${errorText(saveError)}`);
      }
    } finally {
      savingRef.current = false;
      if (mountedRef.current) setBusy("");
    }
  }, [canSaveTarget, saveSessionId, savePath, saveCwd, showNotice]);
  const applyExternalUpdate = (0, import_react.useCallback)(() => {
    const pending = externalUpdateRef.current;
    if (!pending) return;
    autoSaverRef.current?.cancel();
    clearWatcherReloadExpected(saveSessionId, savePath);
    forceRebuildRef.current = true;
    setExternalUpdate(null);
    setReloadToken((value) => value + 1);
  }, [saveSessionId, savePath]);
  const reload = (0, import_react.useCallback)(() => {
    if (rendererRef.current && typeof rendererRef.current.reload === "function") {
      try {
        rendererRef.current.reload();
        return;
      } catch {
      }
    }
    setReloadToken((value) => value + 1);
  }, []);
  const openExternal = (0, import_react.useCallback)(() => {
    if (!mediaUrl) return;
    try {
      window.open(mediaUrl, "_blank", "noopener,noreferrer");
    } catch {
    }
  }, [mediaUrl]);
  const loading = phase === "loading" || phase === "opening" || phase === "converting";
  const failed = phase === "error" && Boolean(error);
  const ready = phase === "ready";
  const canUseEditor = ready && Boolean(managerRef.current);
  const showOwnTitle = props.toolbar !== "host";
  const requestSave = (0, import_react.useCallback)(() => {
    if (!canUseEditor || !managerRef.current || savingRef.current) return;
    if (readOnly) {
      showNotice(t("saveReadOnly"));
      return;
    }
    if (!hasSaveTarget) {
      showNotice(t("saveNoTarget"));
      return;
    }
    if (!nativeSaveable) {
      showNotice(t("saveUnsupported"));
      return;
    }
    autoSaverRef.current?.cancel();
    void saveToFile(false);
  }, [canUseEditor, readOnly, hasSaveTarget, nativeSaveable, saveToFile, showNotice]);
  const autoSave = (0, import_react.useCallback)(() => {
    if (ignoreChangesUntilRef.current > Date.now()) return;
    if (!canUseEditor || !managerRef.current || savingRef.current) return;
    if (readOnly || !hasSaveTarget || !nativeSaveable) return;
    void saveToFile(true);
  }, [canUseEditor, readOnly, hasSaveTarget, nativeSaveable, saveToFile]);
  autoSaveRef.current = autoSave;
  saveRef.current = requestSave;
  const loadingLabel = phase === "opening" ? t("openingDocument") : phase === "converting" ? t("converting") : t("loadingEngine");
  return /* @__PURE__ */ (0, import_jsx_runtime.jsxs)(
    "div",
    {
      className: "dshoo-root",
      title: canSaveTarget ? `${fileName} — ${saveShortcutLabel()} ${t("save")}` : fileName,
      onKeyDown: interceptSaveKey,
      children: [
        /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-toolbar", role: "toolbar", "aria-label": "OnlyOffice", children: [
          showOwnTitle ? /* @__PURE__ */ (0, import_jsx_runtime.jsx)("span", { className: "dshoo-title", title: fileName, children: fileName }) : null,
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)("span", { className: "dshoo-badge", children: busy === "save" ? t("saving") : readOnly ? t("readOnly") : t("editable") }),
          /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("span", { className: "dshoo-hint-inline", title: t("saveHint"), children: [
            saveShortcutLabel(),
            " ",
            t("save")
          ] }),
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)("span", { className: "dshoo-spacer" }),
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)(
            "button",
            {
              type: "button",
              className: "dshoo-btn",
              "aria-pressed": readOnly,
              title: t("toggleReadOnly"),
              disabled: !canUseEditor || busy !== "",
              onClick: toggleReadOnly,
              children: readOnly ? t("readOnly") : t("editable")
            }
          ),
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)(
            "button",
            {
              type: "button",
              className: "dshoo-btn",
              disabled: !canUseEditor || busy !== "",
              onClick: exportDocument,
              title: t("export"),
              children: busy === "export" ? t("exporting") : t("export")
            }
          )
        ] }),
        externalUpdate ? /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-notice dshoo-notice-action", role: "status", children: [
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)("span", { children: t("externalChanged") }),
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)("button", { type: "button", className: "dshoo-btn", onClick: applyExternalUpdate, disabled: busy !== "", children: t("externalReload") })
        ] }) : null,
        notice ? /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-notice", role: "status", children: [
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)("span", { children: notice }),
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)("button", { type: "button", onClick: () => setNotice(""), title: t("dismiss"), children: "×" })
        ] }) : null,
        /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-stage", children: [
          /* @__PURE__ */ (0, import_jsx_runtime.jsx)(
            "div",
            {
              className: "onlyoffice-container dshoo-frame",
              "data-onlyoffice-container-id": containerId,
              children: /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { id: containerId, style: { position: "absolute", top: 0, right: 0, bottom: 0, left: 0 } })
            }
          ),
          loading ? /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-state", role: "status", children: [
            /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { className: "dshoo-spinner", "aria-hidden": "true" }),
            /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { className: "dshoo-state-title", children: loadingLabel }),
            slowHint ? /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { className: "dshoo-hint", children: t("slowHint") }) : null
          ] }) : null,
          failed ? /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-state", role: "alert", children: [
            /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { className: "dshoo-state-title", "data-tone": "error", children: t("loadFailed") }),
            /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { className: "dshoo-hint", children: stageMessage(error.stage, error.message) }),
            error.detail ? /* @__PURE__ */ (0, import_jsx_runtime.jsx)("pre", { className: "dshoo-detail", children: error.detail }) : null,
            /* @__PURE__ */ (0, import_jsx_runtime.jsx)("div", { className: "dshoo-hint", children: stageHint(error.stage) }),
            /* @__PURE__ */ (0, import_jsx_runtime.jsxs)("div", { className: "dshoo-state-actions", children: [
              /* @__PURE__ */ (0, import_jsx_runtime.jsx)("button", { type: "button", className: "dshoo-btn", onClick: reload, children: t("retry") }),
              /* @__PURE__ */ (0, import_jsx_runtime.jsx)("button", { type: "button", className: "dshoo-btn", onClick: openExternal, disabled: !mediaUrl, children: t("openExternal") })
            ] })
          ] }) : null
        ] })
      ]
    }
  );
}

// src/client/icons.jsx
var import_jsx_runtime2 = require("react/jsx-runtime");
var COMMON = {
  xmlns: "http://www.w3.org/2000/svg",
  viewBox: "0 0 16 16",
  fill: "none",
  stroke: "currentColor",
  strokeWidth: 1.2,
  strokeLinecap: "round",
  strokeLinejoin: "round",
  "aria-hidden": true,
  focusable: "false"
};
function Svg({ size = 16, children }) {
  return /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("svg", { width: size, height: size, ...COMMON, children });
}
function IconDocx({ size = 16 } = {}) {
  return /* @__PURE__ */ (0, import_jsx_runtime2.jsxs)(Svg, { size, children: [
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M3.5 1.5h6l3 3v9a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M9.5 1.5v3h3" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M5.2 8h5.6M5.2 10.4h5.6M5.2 12.8h3.4" })
  ] });
}
function IconXlsx({ size = 16 } = {}) {
  return /* @__PURE__ */ (0, import_jsx_runtime2.jsxs)(Svg, { size, children: [
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M3.5 1.5h6l3 3v9a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M9.5 1.5v3h3" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M4.8 7.4h6.4v6H4.8z" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M4.8 9.7h6.4M7 7.4v6" })
  ] });
}
function IconPptx({ size = 16 } = {}) {
  return /* @__PURE__ */ (0, import_jsx_runtime2.jsxs)(Svg, { size, children: [
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M3.5 1.5h6l3 3v9a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M9.5 1.5v3h3" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M4.8 7.2h6.4v4.4H4.8z" }),
    /* @__PURE__ */ (0, import_jsx_runtime2.jsx)("path", { d: "M3.4 13.4h9.2" })
  ] });
}

// src/client/styles.js
var STYLE_ATTRIBUTE = "data-dsh-onlyoffice-styles";
var CSS = `
.dshoo-root {
  position: relative;
  display: flex;
  flex-direction: column;
  width: 100%;
  height: 100%;
  min-width: 0;
  min-height: 0;
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-primary);
  font: inherit;
  font-size: 12px;
  line-height: 1.4;
}

.dshoo-toolbar {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 6px;
  min-height: 34px;
  padding: 4px 8px;
  box-sizing: border-box;
  background: var(--dsw-alias-bg-layer-1);
  border-bottom: 1px solid var(--dsw-alias-border-l1);
  overflow-x: auto;
  scrollbar-width: none;
}
.dshoo-toolbar::-webkit-scrollbar { display: none; }

.dshoo-title {
  flex: 0 1 auto;
  min-width: 0;
  max-width: 42%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--dsw-alias-label-secondary);
  font-size: 12px;
}

.dshoo-spacer { flex: 1 1 auto; }

/* The save action lives on Cmd/Ctrl+S; this is its discoverability hint. */
.dshoo-hint-inline {
  flex: 0 0 auto;
  padding: 0 4px;
  color: var(--dsw-alias-label-secondary);
  font-size: 10px;
  line-height: 16px;
  white-space: nowrap;
  opacity: 0.75;
  user-select: none;
}

.dshoo-badge {
  flex: 0 0 auto;
  padding: 0 6px;
  border: 1px solid var(--dsw-alias-border-l2);
  border-radius: 999px;
  color: var(--dsw-alias-label-secondary);
  font-size: 10px;
  line-height: 16px;
  white-space: nowrap;
}

.dshoo-btn {
  flex: 0 0 auto;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  height: 24px;
  padding: 0 8px;
  box-sizing: border-box;
  border: 1px solid var(--dsw-alias-border-l2);
  border-radius: 6px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 12px;
  line-height: 1;
  white-space: nowrap;
  cursor: pointer;
  user-select: none;
}
.dshoo-btn:hover:not(:disabled) {
  background: var(--dsw-alias-bg-layer-2);
  color: var(--dsw-alias-label-primary);
}
.dshoo-btn:focus-visible {
  outline: 2px solid var(--dsw-alias-brand-primary);
  outline-offset: 1px;
}
.dshoo-btn:disabled {
  opacity: 0.5;
  cursor: default;
}
.dshoo-btn[aria-pressed="true"] {
  border-color: transparent;
  background: var(--dsw-alias-brand-primary);
  color: var(--dsw-alias-bg-base);
}

.dshoo-notice {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 10px;
  border-bottom: 1px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-layer-2);
  color: var(--dsw-alias-state-warn-primary);
  font-size: 11px;
}
.dshoo-notice > span { flex: 1 1 auto; min-width: 0; }

/* The external-change banner has its own action button. */
.dshoo-notice-action {
  color: var(--dsw-alias-label-primary);
}
.dshoo-notice-action .dshoo-btn {
  height: 22px;
  padding: 0 8px;
  font-size: 11px;
}
.dshoo-notice button {
  flex: 0 0 auto;
  border: 0;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  cursor: pointer;
  padding: 0 2px;
}
.dshoo-notice button:hover { color: var(--dsw-alias-label-primary); }

.dshoo-stage {
  position: relative;
  flex: 1 1 auto;
  min-width: 0;
  min-height: 0;
  background: var(--dsw-alias-bg-base);
}

.dshoo-frame {
  position: absolute;
  inset: 0;
  overflow: hidden;
}

.dshoo-state {
  position: absolute;
  inset: 0;
  z-index: 2;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 20px;
  box-sizing: border-box;
  text-align: center;
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-secondary);
}

.dshoo-state-title {
  color: var(--dsw-alias-label-primary);
  font-size: 13px;
  font-weight: 600;
}
.dshoo-state-title[data-tone="error"] { color: var(--dsw-alias-state-error-primary); }

.dshoo-hint {
  max-width: 46ch;
  color: var(--dsw-alias-label-secondary);
  font-size: 11px;
  overflow-wrap: anywhere;
}

.dshoo-state-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  justify-content: center;
  margin-top: 2px;
}

.dshoo-detail {
  max-width: 100%;
  max-height: 120px;
  overflow: auto;
  margin: 0;
  padding: 6px 8px;
  border: 1px solid var(--dsw-alias-border-l1);
  border-radius: 6px;
  background: var(--dsw-alias-bg-layer-1);
  color: var(--dsw-alias-label-secondary);
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 10px;
  text-align: left;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.dshoo-spinner {
  width: 18px;
  height: 18px;
  border: 2px solid var(--dsw-alias-border-l2);
  border-top-color: var(--dsw-alias-brand-primary);
  border-radius: 50%;
  animation: dshoo-spin 0.9s linear infinite;
}

@keyframes dshoo-spin {
  to { transform: rotate(360deg); }
}

@media (prefers-reduced-motion: reduce) {
  .dshoo-spinner { animation-duration: 2.4s; }
}
`;
function installStyles(doc = typeof document === "undefined" ? void 0 : document) {
  if (!doc || !doc.head) return () => {
  };
  if (typeof doc.querySelector === "function" && doc.querySelector(`style[${STYLE_ATTRIBUTE}]`)) {
    return () => {
    };
  }
  const style = doc.createElement("style");
  style.setAttribute(STYLE_ATTRIBUTE, "dsh-onlyoffice");
  style.textContent = CSS;
  doc.head.appendChild(style);
  return () => {
    try {
      style.remove();
    } catch {
    }
  };
}

// src/client/native.jsx
var import_react2 = require("react");

// src/client/effect.js
function registerEffect(ctx, factory, label) {
  try {
    if (ctx && typeof ctx.effect === "function") {
      ctx.effect(factory, label);
      return;
    }
  } catch (error) {
    console.warn(`[dsh-onlyoffice] ${label} failed`, error);
    return;
  }
  try {
    factory();
  } catch (error) {
    console.warn(`[dsh-onlyoffice] ${label} failed`, error);
  }
}

// src/client/native.jsx
var NATIVE_DOCUMENT_ID = "dsh-onlyoffice:document";
var NATIVE_EXTENSIONS = Object.freeze([
  "docx",
  "doc",
  "odt",
  "rtf",
  "xlsx",
  "xls",
  "ods",
  "csv",
  "pptx",
  "ppt",
  "odp"
]);
var EXT_TO_KIND = Object.freeze({
  docx: "docx",
  doc: "docx",
  odt: "docx",
  rtf: "docx",
  xlsx: "xlsx",
  xls: "xlsx",
  ods: "xlsx",
  csv: "xlsx",
  pptx: "pptx",
  ppt: "pptx",
  odp: "pptx"
});
function extensionOf(name) {
  const match = /\.([A-Za-z0-9]{1,8})$/.exec(String(name || ""));
  return match ? match[1].toLowerCase() : "";
}
function baseNameOf(filePath) {
  const parts = String(filePath || "").split(/[\\/]/);
  return parts[parts.length - 1] || "";
}
function parseFileResourceAddress(address) {
  const empty = { scope: "", sessionId: "", path: "" };
  const text = String(address || "");
  const prefix = "dsh-resource://file/";
  if (!text.startsWith(prefix)) return empty;
  const cut = text.search(/[?#]/);
  const body = text.slice(prefix.length, cut === -1 ? void 0 : cut);
  const [scope, ...rest] = body.split("/");
  if (scope === "session") {
    const [rawSession, ...segments] = rest;
    if (!rawSession || segments.length === 0) return empty;
    try {
      return {
        scope: "session",
        sessionId: decodeURIComponent(rawSession),
        path: segments.map(decodeURIComponent).join("/")
      };
    } catch {
      return empty;
    }
  }
  if (scope === "absolute") {
    const unc = rest[0] === "" && rest.length > 1;
    const segments = (unc ? rest.slice(1) : rest).map((segment) => {
      try {
        return decodeURIComponent(segment);
      } catch {
        return segment;
      }
    });
    if (segments.length === 0 || segments[0] === "") return empty;
    const windowsStyle = /^[A-Za-z]:$/.test(segments[0]);
    return {
      scope: "absolute",
      sessionId: "",
      path: unc ? `//${segments.join("/")}` : windowsStyle ? segments.join("/") : `/${segments.join("/")}`
    };
  }
  return empty;
}
function NativeDocumentBody(props) {
  const address = props && props.resourceAddress ? props.resourceAddress : "";
  const parsedAddress = parseFileResourceAddress(address);
  const filePath = parsedAddress.path;
  const fileName = baseNameOf(filePath);
  const kind = EXT_TO_KIND[extensionOf(fileName)] || EXT_TO_KIND[extensionOf(filePath)] || "docx";
  const content = props ? props.content : void 0;
  const owner = content && content.kind === "renderer" ? content : void 0;
  const pluginCtx = props ? props.ctx : void 0;
  const renderer = owner ? {
    revision: owner.revision,
    loaded: owner.loaded,
    failed: owner.failed,
    reload: owner.reload,
    load: (signal) => readDocumentFile(pluginCtx, {
      sessionId: parsedAddress.sessionId,
      path: filePath,
      signal
    })
  } : void 0;
  return (0, import_react2.createElement)(OfficeView, {
    ctx: pluginCtx,
    kind,
    // `path` is what the editor shows; `sourcePath` + `sessionId` are what the
    // save-back route writes to (path may be absolute or session-relative).
    path: filePath,
    sourcePath: filePath,
    sessionId: parsedAddress.sessionId,
    title: fileName || `document.${kind}`,
    renderer
  });
}
function installNativeDocumentPreview(ctx) {
  const previews = ctx && ctx.documentPreviews;
  const slots = ctx && ctx.slots;
  if (!previews || typeof previews.register !== "function" || !slots || typeof slots.inject !== "function" || typeof slots.register !== "function") {
    return false;
  }
  registerEffect(
    ctx,
    () => previews.register({
      id: NATIVE_DOCUMENT_ID,
      extensions: NATIVE_EXTENSIONS,
      binaryExtensions: NATIVE_EXTENSIONS,
      priority: "extension",
      title: () => t("viewerOnlyOffice"),
      loading: "renderer"
    }),
    "dsh-onlyoffice: document preview implementation"
  );
  const Body = (props) => (0, import_react2.createElement)(NativeDocumentBody, Object.assign({}, props, { ctx }));
  registerEffect(
    ctx,
    () => slots.inject(
      "sidebar.right.tab.document",
      () => slots.register(
        { name: "sidebar.right.tab.document", key: NATIVE_DOCUMENT_ID },
        Body
      )
    ),
    "dsh-onlyoffice: document preview body"
  );
  return true;
}

// src/client/index.jsx
var inject = ["documentPreviews", "slots", "locale"];
var VIEWER_EXTS = {
  docx: ["docx", "doc", "odt", "rtf"],
  xlsx: ["xlsx", "xls", "ods", "csv"],
  pptx: ["pptx", "ppt", "odp"]
};
function viewerComponent(kind) {
  const Component = function DshOnlyOfficeFileView(props) {
    return (0, import_react3.createElement)(OfficeView, Object.assign({}, props, { kind }));
  };
  Component.displayName = `DshOnlyOfficeFileView(${kind})`;
  return Component;
}
var COMPONENTS = {
  docx: viewerComponent("docx"),
  xlsx: viewerComponent("xlsx"),
  pptx: viewerComponent("pptx")
};
var ICONS = {
  docx: IconDocx,
  xlsx: IconXlsx,
  pptx: IconPptx
};
var VIEWER_TITLES = {
  docx: () => t("viewerDocx"),
  xlsx: () => t("viewerXlsx"),
  pptx: () => t("viewerPptx")
};
var viewerDescriptors = ["docx", "xlsx", "pptx"].map((kind) => ({
  id: `dsh-onlyoffice:${kind}`,
  title: VIEWER_TITLES[kind],
  icon: (size) => (0, import_react3.createElement)(ICONS[kind], { size }),
  exts: VIEWER_EXTS[kind],
  priority: 10,
  fetchStrategy: "mediaUrl",
  component: COMPONENTS[kind]
}));
function readSidebarService(ctx) {
  try {
    const viaGet = ctx && typeof ctx.get === "function" ? ctx.get("betterSidebar") : void 0;
    if (viaGet && typeof viaGet.registerFileViewer === "function") return viaGet;
  } catch {
  }
  try {
    const direct = ctx && ctx.betterSidebar;
    if (direct && typeof direct.registerFileViewer === "function") return direct;
  } catch {
  }
  try {
    const viaGet = ctx && typeof ctx.get === "function" ? ctx.get("betterSidebar") : void 0;
    if (viaGet && typeof viaGet.registerFileViewer === "function") return viaGet;
  } catch {
  }
  return void 0;
}
var BETTER_SIDEBAR_POLL_MS = 400;
var BETTER_SIDEBAR_POLL_TICKS = 75;
function installBetterSidebarViewers(ctx) {
  let attached = false;
  const attach = (service) => {
    if (attached) return true;
    if (!service || typeof service.registerFileViewer !== "function") return false;
    attached = true;
    for (const descriptor of viewerDescriptors) {
      registerEffect(
        ctx,
        () => service.registerFileViewer(descriptor),
        `dsh-onlyoffice: viewer ${descriptor.id}`
      );
    }
    return true;
  };
  if (attach(readSidebarService(ctx))) return true;
  let canInject = false;
  try {
    canInject = Boolean(ctx) && typeof ctx.inject === "function";
  } catch {
    canInject = false;
  }
  if (canInject) {
    try {
      ctx.inject(["betterSidebar"], (scoped) => {
        attach(readSidebarService(scoped));
      });
    } catch (error) {
      console.warn("[dsh-onlyoffice] optional betterSidebar injection failed", error);
    }
  }
  registerEffect(
    ctx,
    () => {
      let ticks = 0;
      const timer = setInterval(() => {
        ticks += 1;
        if (attach(readSidebarService(ctx)) || ticks >= BETTER_SIDEBAR_POLL_TICKS) {
          clearInterval(timer);
        }
      }, BETTER_SIDEBAR_POLL_MS);
      return () => clearInterval(timer);
    },
    "dsh-onlyoffice: betterSidebar waiter"
  );
  return attached;
}
function apply(ctx) {
  registerEffect(ctx, () => installStyles(), "dsh-onlyoffice: viewer styles");
  registerEffect(
    ctx,
    () => installDiagnostics(async (text) => {
      try {
        await fetch(resolveHostUrl(ctx, "dsh-onlyoffice/diagnostic"), {
          method: "POST",
          credentials: "same-origin",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ text }),
          keepalive: true
        });
      } catch {
      }
    }),
    "dsh-onlyoffice: diagnostics"
  );
  registerEffect(ctx, () => installX2tWorkerFactory(), "dsh-onlyoffice: x2t worker factory");
  registerEffect(ctx, () => attachLocale(ctx), "dsh-onlyoffice: locale tracking");
  const native = installNativeDocumentPreview(ctx);
  const betterSidebar = installBetterSidebarViewers(ctx);
  if (!native && !betterSidebar) {
    console.warn(
      "[dsh-onlyoffice] neither the built-in document preview nor betterSidebar is available; Office viewers were not registered."
    );
  }
}
return module.exports; } });

