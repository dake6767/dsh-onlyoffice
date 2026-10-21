// SPDX-License-Identifier: AGPL-3.0-only
// dsh-onlyoffice — see LICENSE and THIRD_PARTY_NOTICES.md.

/**
 * Viewer icons for the sidebar settings inventory.
 *
 * Hand-drawn 16×16 glyphs using `currentColor` only, so they follow the host
 * theme (the host renders descriptor icons inside its own settings list).
 */

const COMMON = {
  xmlns: 'http://www.w3.org/2000/svg',
  viewBox: '0 0 16 16',
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 1.2,
  strokeLinecap: 'round',
  strokeLinejoin: 'round',
  'aria-hidden': true,
  focusable: 'false',
};

function Svg({ size = 16, children }) {
  return (
    <svg width={size} height={size} {...COMMON}>
      {children}
    </svg>
  );
}

export function IconDocx({ size = 16 } = {}) {
  return (
    <Svg size={size}>
      <path d="M3.5 1.5h6l3 3v9a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z" />
      <path d="M9.5 1.5v3h3" />
      <path d="M5.2 8h5.6M5.2 10.4h5.6M5.2 12.8h3.4" />
    </Svg>
  );
}

export function IconXlsx({ size = 16 } = {}) {
  return (
    <Svg size={size}>
      <path d="M3.5 1.5h6l3 3v9a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z" />
      <path d="M9.5 1.5v3h3" />
      <path d="M4.8 7.4h6.4v6H4.8z" />
      <path d="M4.8 9.7h6.4M7 7.4v6" />
    </Svg>
  );
}

export function IconPptx({ size = 16 } = {}) {
  return (
    <Svg size={size}>
      <path d="M3.5 1.5h6l3 3v9a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-11a1 1 0 0 1 1-1Z" />
      <path d="M9.5 1.5v3h3" />
      <path d="M4.8 7.2h6.4v4.4H4.8z" />
      <path d="M3.4 13.4h9.2" />
    </Svg>
  );
}
