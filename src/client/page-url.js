/**
 * Page-relative route resolution.
 *
 * DSH can be served from three different bases and this plugin's Host routes
 * must follow whichever one is in force:
 *
 *  - origin root (`http://127.0.0.1:53803/`)             → `/dsh-onlyoffice/...`
 *  - a reverse-proxy directory (`https://host/app/<id>/`) → `/app/<id>/dsh-onlyoffice/...`
 *  - the desktop shell (`dsh-app://app/`)                → `dsh-app://app/dsh-onlyoffice/...`
 *
 * Resolving a route against `document.baseURI` keeps the request under the
 * page's own directory. Resolving it against `location.origin` (the old
 * behaviour) dropped a reverse-proxy prefix and produced a 404 for every SDK
 * asset on a sub-path deployment.
 */

/**
 * Resolve a route relative to a page base.
 *
 * @param {string} path route path, with or without a leading slash
 * (`dsh-onlyoffice/status`); leading slashes are stripped on purpose so they
 * cannot escape the page's directory.
 * @param {string} [baseHref] absolute page URL (typically `document.baseURI`)
 * @param {string} [fallbackOrigin] origin used when no usable base exists
 * @returns {string} absolute URL, without a trailing slash
 */
export function resolvePageRelative(path, baseHref, fallbackOrigin) {
  const clean = String(path || '').replace(/^\/+/, '')
  const withSlash = (url) => {
    if (!url.pathname.endsWith('/')) url.pathname += '/'
    return url
  }

  if (typeof baseHref === 'string' && baseHref !== '') {
    try {
      const base = withSlash(new URL(baseHref))
      base.search = ''
      base.hash = ''
      return new URL(clean, base).href.replace(/\/+$/, '')
    } catch {
      /* fall through to the origin form */
    }
  }

  const origin = typeof fallbackOrigin === 'string' ? fallbackOrigin : ''
  if (!origin || origin === 'null' || origin === 'undefined') {
    // Opaque origin (dsh-app://, file:): a root-relative path stays valid and
    // resolves against whatever scheme hosts the page.
    return `/${clean}`.replace(/\/+$/, '')
  }
  return `${origin.replace(/\/+$/, '')}/${clean}`.replace(/\/+$/, '')
}
