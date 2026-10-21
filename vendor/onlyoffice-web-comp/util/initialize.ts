import { STATIC_RESOURCE } from "../const";

let initializePromise: Promise<void> | null = null;
let initializeApiUrl = "";

type DocsApiWindow = Window & {
  DocsAPI?: unknown;
};

function resetLoadedDocsApi(apiUrl: string) {
  if (!initializeApiUrl || initializeApiUrl === apiUrl) {
    return;
  }

  document
    .querySelectorAll<HTMLScriptElement>(
      'script[src*="/web-apps/apps/api/documents/api.js"]',
    )
    .forEach((script) => script.remove());
  document
    .querySelectorAll<HTMLIFrameElement>("iframe[data-onlyoffice-preload]")
    .forEach((iframe) => iframe.remove());

  try {
    delete (window as DocsApiWindow).DocsAPI;
  } catch {
    (window as DocsApiWindow).DocsAPI = undefined;
  }
}

function preloadEditorFrame() {
  if (document.querySelector(`iframe[data-onlyoffice-preload="${STATIC_RESOURCE.onlyoffice.preloadHtml}"]`)) {
    return;
  }

  const iframe = document.createElement("iframe");
  iframe.src = STATIC_RESOURCE.onlyoffice.preloadUrl;
  iframe.dataset.onlyofficePreload = STATIC_RESOURCE.onlyoffice.preloadHtml;
  /**
   * DSH patch: upstream used Tailwind utility classes (`w-0 h-0 hidden …`) that
   * do not exist in the host page, leaving a visible 300×150 preload iframe in
   * the DSH document body. Inline styles keep it invisible without a stylesheet.
   */
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

export async function initializeOnlyOffice() {
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

  initializePromise = new Promise<void>((resolve, reject) => {
    preloadEditorFrame();

    if (window.DocsAPI?.DocEditor) {
      resolve();
      return;
    }

    let script = document.querySelector<HTMLScriptElement>(
      `script[src="${apiUrl}"]`,
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
      { once: true },
    );
  });

  return initializePromise;
}
