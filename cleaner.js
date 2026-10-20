// HTML Cleaner — turns clipboard HTML (Word, Google Docs, web pages) or plain text into
// clean, basic HTML with no styling and minimal attributes. No UI coupling: parses with the
// global DOMParser (native in browsers, provided by jsdom in tests).

const ELEMENT = 1;
const TEXT = 3;
const COMMENT = 8;
const SHOW_COMMENT = 128;

const BLOCK_TAGS = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'hr', 'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'th', 'td']);
const INLINE_TAGS = new Set(['strong', 'em', 'u', 's', 'sub', 'sup', 'code', 'a']);
const ALLOWED_TAGS = new Set([...BLOCK_TAGS, ...INLINE_TAGS, 'br', 'img']);
const HEADING_SEL = 'h1,h2,h3,h4,h5,h6';
const INLINE_SEL = [...INLINE_TAGS].join(',');
// Containers whose children are fixed by structure (list items, rows, cells), never loose inline content.
const STRUCTURAL = new Set(['ul', 'ol', 'table', 'thead', 'tbody', 'tr']);

const ALLOWED_ATTRS = {
  a: ['href', 'target', 'title'],
  img: ['src', 'alt'],
  td: ['colspan', 'rowspan'],
  th: ['colspan', 'rowspan'],
  ol: ['start', 'type'],
};

// Removed together with their content.
const DROP_TAGS = new Set([
  'script', 'style', 'noscript', 'template', 'iframe', 'frame', 'object', 'embed', 'svg', 'canvas',
  'audio', 'video', 'map', 'input', 'select', 'textarea', 'meta', 'link', 'title', 'base',
]);
// Disallowed tags that behave like blocks: they become <p> when they hold only inline content.
const BLOCKISH_TAGS = new Set([
  'div', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'figure', 'figcaption',
  'address', 'center', 'details', 'summary', 'dl', 'dt', 'dd', 'fieldset', 'legend', 'hgroup', 'form',
]);
const RENAMES = { b: 'strong', i: 'em', strike: 's', del: 's', tt: 'code', kbd: 'code', samp: 'code', menu: 'ul', dir: 'ul' };

// ===================== Public API =====================

/** Cleans clipboard HTML. Returns the output and how many images need a real URL. */
export function cleanHtml(html) {
  html = String(html ?? '');
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const body = doc.body;
  const source = detectSource(html);

  if (source === 'Word') extractVmlShapes(body);
  removeJunk(body);
  flattenPre(body);
  if (source === 'Word') cleanWord(body);
  if (source === 'Google Docs') for (const p of body.querySelectorAll('p')) if (hasClass(p, 'title')) rename(p, 'h1');
  convertFootnoteLinks(body, /^#ftnt(_ref)?\d+$/i);
  removePageBreaks(body);
  stripLeaders(body);
  renameTags(body);
  applyInlineStyles(body);
  removeNamespaced(body);
  fixLinks(body);
  fixImages(body);
  unwrapDisallowed(body);
  fixStructure(body);
  stripAttributes(body);

  for (let i = 0; i < 8; i++) {
    const before = body.innerHTML;
    tidy(body);
    if (body.innerHTML === before) break;
  }

  const emptiedImages = [...body.querySelectorAll('img')].filter(img => img.getAttribute('src') === '').length;
  return { html: serialize(body), emptiedImages };
}

/** Cleans a plain-text paste: blank-line-separated chunks become paragraphs, line breaks become <br>. */
export function cleanText(text) {
  const chunks = String(text ?? '').replace(/\r\n?/g, '\n').split(/\n(?:[ \t]*\n)+/);
  const html = chunks
    .map(chunk => chunk.replace(/^\n+|\n+$/g, ''))
    .filter(chunk => chunk.trim())
    .map(chunk => `<p>${chunk.split('\n').map(escapeText).join('<br>')}</p>`)
    .join('');
  return cleanHtml(html);
}

/** Cleans HTML and returns only the output string. */
export function clean(html) {
  return cleanHtml(html).html;
}

/** Names the application the clipboard HTML came from. */
export function detectSource(html) {
  html = String(html ?? '');
  if (/urn:schemas-microsoft-com:office:excel|name=["']?ProgId["']? content=["']?Excel/i.test(html)) return 'Excel';
  if (/urn:schemas-microsoft-com:office:word|content=["']?Microsoft Word|class=["']?Mso/i.test(html)) return 'Word';
  if (/id=["']?docs-internal-guid-/i.test(html)) return 'Google Docs';
  if (/google-sheets-html-origin/i.test(html)) return 'Google Sheets';
  return 'Web page';
}

/**
 * Lists the tags and attributes in `html` that the allowed set (and so the WYSIWYG schema) would drop,
 * e.g. ['<div>', 'style attributes'].
 */
export function listDisallowed(html) {
  const doc = new DOMParser().parseFromString(String(html ?? ''), 'text/html');
  const tags = new Set();
  const attrs = new Set();
  for (const el of doc.querySelectorAll('head > *, body *')) {
    const tag = el.localName;
    if (!ALLOWED_TAGS.has(tag)) tags.add(`<${tag}>`);
    for (const { name } of el.attributes) if (!ALLOWED_ATTRS[tag]?.includes(name)) attrs.add(`${name} attributes`);
  }
  const comments = doc.createTreeWalker(doc, SHOW_COMMENT).nextNode() ? ['comments'] : [];
  return [...tags, ...[...attrs].sort(), ...comments];
}

// ===================== Small DOM helpers =====================

function styleOf(el) {
  const style = {};
  for (const decl of (el.getAttribute('style') || '').split(';')) {
    const i = decl.indexOf(':');
    if (i < 0) continue;
    style[decl.slice(0, i).trim().toLowerCase()] = decl.slice(i + 1).replace(/!important/i, '').trim().toLowerCase();
  }
  return style;
}

// Class matching by hand: DOMParser documents without a doctype are in quirks mode, where selector
// class matching is case-insensitive in some engines and not others.
function hasClass(el, name) {
  const wanted = name.toLowerCase();
  return (el.getAttribute('class') || '').toLowerCase().split(/\s+/).includes(wanted);
}

function classMatch(el, pattern) {
  return (el.getAttribute('class') || '').split(/\s+/).map(c => c.match(pattern)).find(Boolean) || null;
}

function unwrap(el) {
  el.replaceWith(...el.childNodes);
}

function rename(el, tag) {
  const replacement = el.ownerDocument.createElement(tag);
  for (const { name, value } of el.attributes) replacement.setAttribute(name, value);
  replacement.append(...el.childNodes);
  el.replaceWith(replacement);
  return replacement;
}

function wrapChildren(el, tag) {
  const wrapper = el.ownerDocument.createElement(tag);
  wrapper.append(...el.childNodes);
  el.append(wrapper);
}

function isBlock(node) {
  return node.nodeType === ELEMENT && BLOCK_TAGS.has(node.localName);
}

function hasBlockChild(el) {
  return [...el.children].some(isBlock);
}

function hasContent(el) {
  return el.textContent.replace(/[\u200b-\u200d\ufeff]/g, '').trim() !== '' || el.querySelector('img, hr') !== null;
}

function prevNonBlank(node) {
  let prev = node.previousSibling;
  while (prev && prev.nodeType !== ELEMENT && !(prev.nodeType === TEXT && prev.data.trim())) prev = prev.previousSibling;
  return prev;
}

function decodeEntities(doc, text) {
  const t = doc.createElement('textarea');
  t.innerHTML = text;
  return t.value;
}

// ===================== Removal passes =====================

function removeJunk(body) {
  const comments = [];
  const walker = body.ownerDocument.createTreeWalker(body, SHOW_COMMENT);
  while (walker.nextNode()) comments.push(walker.currentNode);
  comments.forEach(c => c.remove());

  for (const el of [...body.querySelectorAll('*')]) {
    if (!el.isConnected) continue;
    const style = styleOf(el);
    if (DROP_TAGS.has(el.localName) || el.localName === 'o:p' || el.hasAttribute('hidden') || style.display === 'none' || style['mso-hide'] === 'all') el.remove();
  }
}

// Code blocks keep only their text; <br> and line-like children become newlines.
function flattenPre(body) {
  for (const pre of [...body.querySelectorAll('pre')]) {
    let text = '';
    const walk = node => {
      for (const child of node.childNodes) {
        if (child.nodeType === TEXT) text += child.data;
        else if (child.nodeType === ELEMENT) {
          if (child.localName === 'br') text += '\n';
          else {
            walk(child);
            if (/^(div|p|li|tr)$/.test(child.localName) && !text.endsWith('\n')) text += '\n';
          }
        }
      }
    };
    walk(pre);
    pre.textContent = text.replace(/\u00a0/g, ' ').replace(/\n+$/, '');
  }
}

function removePageBreaks(body) {
  for (const el of [...body.querySelectorAll('br, hr')]) {
    const style = styleOf(el);
    const breaks = ['page-break-before', 'page-break-after', 'break-before', 'break-after'].some(p => /always|page/.test(style[p] || ''));
    if (breaks || style['mso-break-type']) el.remove();
  }
}

// Table-of-contents entries: dot leaders followed by a page number.
function stripLeaders(body) {
  for (const block of body.querySelectorAll('p')) {
    if (!classMatch(block, /^MsoToc\d$/i)) continue;
    let last = block.lastChild;
    while (last && last.nodeType === ELEMENT) last = last.lastChild;
    if (last?.nodeType !== TEXT) continue;
    const text = last.data.trimEnd();
    const tail = text.slice(-200).match(/[\s\u00a0]*(?:\.[\s\u00a0]?){4,}\d+$/);
    if (tail) last.data = text.slice(0, text.length - tail[0].length);
  }
}

function removeNamespaced(body) {
  for (const el of [...body.querySelectorAll('*')]) if (el.localName.includes(':')) unwrap(el);
}

// ===================== Word =====================

// Word puts drawing objects in VML inside conditional comments. WordArt text only exists there, and
// the shape ids tell picture fallbacks (<img v:shapes>) apart from shape fallbacks, which are dropped.
function extractVmlShapes(body) {
  const doc = body.ownerDocument;
  const walker = doc.createTreeWalker(body, SHOW_COMMENT);
  const comments = [];
  while (walker.nextNode()) comments.push(walker.currentNode);

  const shapeIds = new Set();
  const shapeRe = /<v:(shape|rect|roundrect|oval|line|polyline|arc)\b([^>]*?)(\/>|>([\s\S]*?)<\/v:\1>)/gi;
  for (const comment of comments) {
    for (const [, , attrs, , inner = ''] of comment.data.matchAll(shapeRe)) {
      const id = attrs.match(/\bid=["']?([^"'\s>]+)/i)?.[1];
      const wordArt = inner.match(/<v:textpath\b[^>]*?\bstring=["']([^"']*)["']/i);
      if (wordArt) {
        const holder = doc.createElement('div');
        const p = doc.createElement('p');
        p.textContent = decodeEntities(doc, wordArt[1]);
        holder.append(p);
        comment.before(holder);
      }
      if (id && !/<v:imagedata\b/i.test(inner)) shapeIds.add(id);
    }
  }
  for (const img of [...body.querySelectorAll('img')]) {
    const ids = (img.getAttribute('v:shapes') || '').split(/[\s,]+/);
    if (ids.some(id => shapeIds.has(id))) img.remove();
  }
}

function cleanWord(body) {
  const doc = body.ownerDocument;

  // Comments: anchors, reference marks and the comment list.
  for (const el of [...body.querySelectorAll('*')]) {
    if (!el.isConnected) continue;
    const element = styleOf(el)['mso-element'];
    if (['msocomanchor', 'msocomreference', 'msocommentreference', 'msocomoff', 'msocomtxt'].some(c => hasClass(el, c)) || element === 'comment-list' || element === 'comment') el.remove();
  }

  // Tracked changes: keep insertions, drop deletions.
  for (const el of [...body.querySelectorAll('del, .msoDel, .msodel')]) el.remove();
  for (const el of [...body.querySelectorAll('ins')]) unwrap(el);

  // Footnotes and endnotes: marks become <sup>, the note lists move to the end without separators.
  convertFootnoteLinks(body, /^#_(ftn|edn)(ref)?\d+$/i);
  for (const list of [...body.querySelectorAll('div')].filter(d => /^(footnote|endnote)-list$/.test(styleOf(d)['mso-element'] || ''))) {
    for (const child of [...list.children]) if (child.localName === 'hr' || child.localName === 'br') child.remove();
    body.append(list);
  }

  // Text boxes: their paragraphs replace the floating layout span they sit in.
  for (const span of [...body.querySelectorAll('span')]) {
    if (!span.isConnected || styleOf(span)['mso-ignore'] !== 'vglayout') continue;
    let outer = span;
    for (let up = span.parentElement; up && up !== body; up = up.parentElement) if (styleOf(up)['mso-ignore'] === 'vglayout') outer = up;
    const holder = doc.createElement('div');
    for (const box of outer.querySelectorAll('*')) if (box.hasAttribute('v:shape')) holder.append(...box.childNodes);
    holder.append(...outer.querySelectorAll('img'));
    outer.replaceWith(holder);
  }

  // Table of contents: drop the tab leader and everything after it (the page number).
  for (const p of [...body.querySelectorAll('p')].filter(p => classMatch(p, /^MsoToc\d$/i))) {
    const tabs = [...p.querySelectorAll('span')].filter(s => styleOf(s)['mso-tab-count']);
    if (!tabs.length) continue;
    const range = doc.createRange();
    range.setStartBefore(tabs.at(-1));
    range.setEndAfter(p.lastChild);
    range.deleteContents();
  }

  for (const br of [...body.querySelectorAll('br[clear]')]) br.remove();

  rebuildWordLists(body);

  for (const p of [...body.querySelectorAll('p')]) {
    if (classMatch(p, /^MsoTitle$/i)) rename(p, 'h1');
    else {
      const level = classMatch(p, /^MsoHeading([1-6])$/i)?.[1];
      if (level) rename(p, `h${level}`);
    }
  }
}

function convertFootnoteLinks(body, hrefPattern) {
  for (const a of [...body.querySelectorAll('a')]) {
    const style = styleOf(a);
    if (!hrefPattern.test(a.getAttribute('href') || '') && !style['mso-footnote-id'] && !style['mso-endnote-id']) continue;
    const sup = a.ownerDocument.createElement('sup');
    sup.textContent = a.textContent.replace(/[[\]\s\u00a0]/g, '');
    a.replaceWith(sup);
  }
}

// Word's lists are paragraphs with an `mso-list: l0 level1 lfo1` style and a literal marker.
function wordListInfo(p) {
  const doc = p.ownerDocument;
  const m = (styleOf(p)['mso-list'] || '').match(/l(\d+)\s+level(\d+)\s+lfo(\d+)/);
  if (m) {
    const ignore = [...p.querySelectorAll('span')].find(s => styleOf(s)['mso-list'] === 'ignore');
    const marker = ignore ? ignore.textContent.replace(/[\s\u00a0]+/g, '') : '';
    ignore?.remove();
    return { list: `${m[1]}/${m[3]}`, level: +m[2], marker };
  }
  // A paragraph that starts with a literal bullet character.
  const walker = doc.createTreeWalker(p, 4);
  let first = walker.nextNode();
  while (first && !first.data.trim()) first = walker.nextNode();
  const bullet = first?.data.match(/^[\s\u00a0]*([•·▪◦‣⁃●○■])[\s\u00a0]+/);
  if (!bullet) return null;
  first.data = first.data.slice(bullet[0].length);
  return { list: 'literal', level: 1, marker: bullet[1] };
}

function romanValue(s) {
  const values = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  const digits = [...s.toLowerCase()].map(c => values[c]);
  return digits.reduce((sum, v, i) => sum + (v < (digits[i + 1] || 0) ? -v : v), 0);
}

function parseMarker(marker) {
  let m = marker.match(/^(?:\d+\.)+(\d+)\.?$/) || marker.match(/^\(?(\d+)[.)]?$/);
  if (m) return { tag: 'ol', type: '1', value: +m[1] };
  m = marker.match(/^\(?([a-zA-Z]+)[.)]$/);
  if (m) {
    const s = m[1];
    const lower = s === s.toLowerCase();
    if (/^(i|[ivxlcdm]{2,})$/i.test(s)) return { tag: 'ol', type: lower ? 'i' : 'I', value: romanValue(s) };
    if (/^(.)\1*$/.test(s)) return { tag: 'ol', type: lower ? 'a' : 'A', value: (s.length - 1) * 26 + s.toLowerCase().charCodeAt(0) - 96 };
  }
  return { tag: 'ul' };
}

function rebuildWordLists(body) {
  const doc = body.ownerDocument;
  let run = null; // { list, last: top-level list element, stack: [{ level, el, tag, li }] }

  for (const p of [...body.querySelectorAll('p')]) {
    if (!p.isConnected) continue;
    const info = wordListInfo(p);
    if (!info) continue;
    const fmt = parseMarker(info.marker);
    if (!run || run.list !== info.list || prevNonBlank(p) !== run.last) run = { list: info.list, last: null, stack: [] };

    const { stack } = run;
    while (stack.length && stack.at(-1).level > info.level) stack.pop();
    if (stack.length && stack.at(-1).level === info.level && stack.at(-1).tag !== fmt.tag) stack.pop();

    if (!stack.length || stack.at(-1).level < info.level) {
      const list = doc.createElement(fmt.tag);
      if (fmt.tag === 'ol') {
        if (fmt.value !== 1) list.setAttribute('start', String(fmt.value));
        if (fmt.type !== '1') list.setAttribute('type', fmt.type);
      }
      if (stack.length) stack.at(-1).li.append(list);
      else {
        if (run.last) run.last.after(list);
        else p.before(list);
        run.last = list;
      }
      stack.push({ level: info.level, el: list, tag: fmt.tag, li: null });
    }

    const li = doc.createElement('li');
    if (p.hasAttribute('style')) li.setAttribute('style', p.getAttribute('style'));
    li.append(...p.childNodes);
    const top = stack.at(-1);
    top.el.append(li);
    top.li = li;
    p.remove();
  }
}

// ===================== Formatting =====================

function renameTags(body) {
  for (const el of [...body.querySelectorAll(Object.keys(RENAMES).join(','))]) rename(el, RENAMES[el.localName]);
}

// Bold and italic expressed as inline styles become real tags before styles are stripped.
function applyInlineStyles(body) {
  for (const el of [...body.querySelectorAll('*')]) {
    if (!el.isConnected || el.localName === 'img' || el.localName === 'br' || el.localName === 'hr') continue;
    const style = styleOf(el);
    const weight = style['font-weight'];
    const fontStyle = style['font-style'];
    if (el.localName === 'strong' && /^(normal|lighter|[1-5]00)$/.test(weight || '')) { unwrap(el); continue; }
    if (el.localName === 'em' && fontStyle === 'normal') { unwrap(el); continue; }
    if (/^(bold|bolder|[6-9]00)$/.test(weight || '') && !el.closest('strong')) wrapChildren(el, 'strong');
    if (/^(italic|oblique)/.test(fontStyle || '') && !el.closest('em')) wrapChildren(el, 'em');
  }
}

function fixLinks(body) {
  for (const a of [...body.querySelectorAll('a')]) {
    const href = (a.getAttribute('href') || '').replace(/[\s\u0000-\u001f]+/g, '');
    if (!href || href === '#' || /^(javascript|vbscript|data):/i.test(href)) unwrap(a);
  }
}

function fixImages(body) {
  for (const img of body.querySelectorAll('img')) {
    const src = (img.getAttribute('src') || '').trim();
    const hasScheme = /^[a-z][a-z0-9+.-]*:/i.test(src);
    img.setAttribute('src', hasScheme && !/^https?:/i.test(src) ? '' : src);
    if (!img.hasAttribute('alt')) img.setAttribute('alt', '');
  }
}

// ===================== Structure =====================

function unwrapDisallowed(body) {
  const doc = body.ownerDocument;
  for (const el of [...body.querySelectorAll('*')].reverse()) {
    const tag = el.localName;
    if (ALLOWED_TAGS.has(tag)) continue;
    if (tag === 'caption') {
      const table = el.closest('table');
      if (table) table.before(rename(el, 'p'));
      else rename(el, 'p');
    } else if (tag === 'colgroup' || tag === 'col') {
      el.remove();
    } else if (tag === 'tfoot') {
      const table = el.closest('table');
      let tbody = [...table.children].filter(c => c.localName === 'tbody').at(-1);
      if (!tbody) table.append((tbody = doc.createElement('tbody')));
      tbody.append(...el.children);
      el.remove();
    } else if (BLOCKISH_TAGS.has(tag)) {
      if (el.querySelector([...BLOCK_TAGS].join(','))) unwrap(el);
      else rename(el, 'p');
    } else {
      unwrap(el);
    }
  }
}

function fixStructure(body) {
  const doc = body.ownerDocument;

  // Inline elements wrapping blocks: push the inline element down into each block.
  for (let found = true; found;) {
    found = false;
    for (const el of body.querySelectorAll(INLINE_SEL)) {
      if (hasBlockChild(el)) { pushInlineDown(el); found = true; break; }
    }
  }

  // Paragraphs and headings that ended up holding blocks are split around them.
  for (const el of [...body.querySelectorAll(`p,${HEADING_SEL}`)].reverse()) {
    if (!hasBlockChild(el)) continue;
    let run = null;
    for (const child of [...el.childNodes]) {
      if (isBlock(child)) { run = null; el.before(child); }
      else {
        if (!run) el.before((run = el.cloneNode(false)));
        run.append(child);
      }
    }
    el.remove();
  }

  // Lists hold only list items; a nested list joins the item before it.
  for (const list of body.querySelectorAll('ul, ol')) {
    let loose = null;
    for (const child of [...list.childNodes]) {
      if (child.nodeType === TEXT && !child.data.trim()) { child.remove(); continue; }
      if (child.nodeType === ELEMENT && child.localName === 'li') { loose = null; continue; }
      if (child.nodeType === ELEMENT && (child.localName === 'ul' || child.localName === 'ol')) {
        const prev = child.previousElementSibling;
        if (prev?.localName === 'li') { prev.append(child); loose = null; continue; }
      }
      if (child.nodeType !== ELEMENT && child.nodeType !== TEXT) { child.remove(); continue; }
      if (!loose) child.before((loose = doc.createElement('li')));
      loose.append(child);
    }
  }
  for (const li of [...body.querySelectorAll('li')]) {
    if (!/^(ul|ol)$/.test(li.parentElement?.localName)) rename(li, 'p');
  }

  // Table sections and rows hold no loose text.
  for (const el of body.querySelectorAll('table, thead, tbody, tr')) {
    for (const child of [...el.childNodes]) if (child.nodeType !== ELEMENT) child.remove();
  }

  // A single paragraph in a list item or cell is unwrapped.
  for (const cell of body.querySelectorAll('li, td, th')) {
    const paragraphs = [...cell.children].filter(c => c.localName === 'p');
    if (paragraphs.length !== 1) continue;
    const others = [...cell.childNodes].filter(c => c !== paragraphs[0]);
    const onlyLists = others.every(c =>
      (c.nodeType === TEXT && !c.data.trim()) || c.nodeType === COMMENT ||
      (cell.localName === 'li' && c.nodeType === ELEMENT && /^(ul|ol)$/.test(c.localName)));
    if (onlyLists && cell.firstElementChild === paragraphs[0]) unwrap(paragraphs[0]);
  }

  // Loose inline content at the top level or in a blockquote goes into paragraphs.
  for (const container of [body, ...body.querySelectorAll('blockquote')]) wrapLooseInline(container);
}

function pushInlineDown(el) {
  let run = null;
  for (const child of [...el.childNodes]) {
    if (isBlock(child)) {
      run = null;
      wrapInside(child, el);
      el.before(child);
    } else {
      if (!run) el.before((run = el.cloneNode(false)));
      run.append(child);
    }
  }
  el.remove();
}

function wrapInside(block, inline) {
  if (block.localName === 'hr' || block.localName === 'pre') return;
  if (!STRUCTURAL.has(block.localName) && !hasBlockChild(block)) {
    const wrapper = inline.cloneNode(false);
    wrapper.append(...block.childNodes);
    block.append(wrapper);
    return;
  }
  let run = null;
  for (const child of [...block.childNodes]) {
    if (isBlock(child)) { run = null; wrapInside(child, inline); }
    else if (!STRUCTURAL.has(block.localName)) {
      if (!run) child.before((run = inline.cloneNode(false)));
      run.append(child);
    }
  }
}

function wrapLooseInline(container) {
  const doc = container.ownerDocument;
  let run = [];
  const flush = () => {
    const meaningful = run.some(n => (n.nodeType === TEXT && n.data.trim()) || (n.nodeType === ELEMENT && n.localName !== 'br' && hasContent(n)) || n.localName === 'img');
    if (meaningful) {
      const p = doc.createElement('p');
      run[0].before(p);
      p.append(...run);
    } else run.forEach(n => n.remove());
    run = [];
  };
  for (const child of [...container.childNodes]) {
    if (isBlock(child)) flush();
    else run.push(child);
  }
  flush();
}

function stripAttributes(body) {
  for (const el of body.querySelectorAll('*')) {
    const tag = el.localName;
    const kept = [];
    const get = name => (el.getAttribute(name) || '').trim();
    if (tag === 'a') {
      kept.push(['href', get('href')]);
      if (get('target')) kept.push(['target', get('target')]);
      if (get('title')) kept.push(['title', get('title')]);
    } else if (tag === 'img') {
      kept.push(['src', get('src')], ['alt', el.getAttribute('alt') || '']);
    } else if (tag === 'td' || tag === 'th') {
      for (const name of ['colspan', 'rowspan']) if (parseInt(get(name), 10) > 1) kept.push([name, String(parseInt(get(name), 10))]);
    } else if (tag === 'ol') {
      const start = parseInt(get('start'), 10);
      if (Number.isFinite(start) && start !== 1) kept.push(['start', String(start)]);
      if (/^[aAiI]$/.test(get('type'))) kept.push(['type', get('type')]);
    }
    for (const { name } of [...el.attributes]) el.removeAttribute(name);
    for (const [name, value] of kept) el.setAttribute(name, value);
  }
}

// ===================== Tidy (repeated until stable) =====================

function tidy(body) {
  normalizeWhitespace(body);

  // Inline elements with no visible content are unwrapped (whitespace and <br> survive).
  for (const el of [...body.querySelectorAll(INLINE_SEL)].reverse()) {
    if (!el.textContent.trim() && !el.querySelector('img')) unwrap(el);
  }

  // The same inline tag nested inside itself adds nothing.
  for (const el of [...body.querySelectorAll(INLINE_SEL)]) {
    if (el.parentElement?.closest(el.localName)) unwrap(el);
  }

  mergeAdjacentInline(body);

  // Headings are already bold.
  for (const h of body.querySelectorAll(HEADING_SEL)) {
    const kids = [...h.childNodes].filter(n => !(n.nodeType === TEXT && !n.data.trim()));
    if (kids.length === 1 && kids[0].localName === 'strong') unwrap(kids[0]);
  }

  // Empty blocks.
  for (const el of [...body.querySelectorAll('p, h1, h2, h3, h4, h5, h6, blockquote, pre, li')].reverse()) {
    if (!hasContent(el)) el.remove();
  }
  for (const el of [...body.querySelectorAll('ul, ol')].reverse()) if (!el.querySelector('li')) el.remove();
  for (const el of [...body.querySelectorAll('tr')]) if (!el.querySelector('td, th')) el.remove();
  for (const el of [...body.querySelectorAll('thead, tbody')]) if (!el.querySelector('tr')) el.remove();
  for (const el of [...body.querySelectorAll('table')]) if (!el.querySelector('tr')) el.remove();

  for (const t of textNodes(body)) if (!t.data) t.remove();
}

function textNodes(root) {
  const nodes = [];
  const walker = root.ownerDocument.createTreeWalker(root, 4);
  while (walker.nextNode()) nodes.push(walker.currentNode);
  return nodes;
}

// Collapses whitespace like a browser would, trims each block's lines, and drops trailing <br>.
function normalizeWhitespace(container) {
  let line = [];
  let atLineStart = true;
  let lastSpace = false;

  const trimEnd = dropBreaks => {
    while (line.length) {
      const node = line.at(-1);
      if (node.nodeType === TEXT) {
        node.data = node.data.replace(/ +$/, '');
        if (node.data) break;
      } else if (node.localName === 'br' && dropBreaks) node.remove();
      else break;
      line.pop();
    }
  };
  const endBlock = () => {
    trimEnd(true);
    line = [];
    atLineStart = true;
    lastSpace = false;
  };
  const visit = node => {
    for (const child of [...node.childNodes]) {
      if (child.nodeType === TEXT) {
        let data = child.data.replace(/\s+/g, ' ');
        if (atLineStart || lastSpace) data = data.replace(/^ /, '');
        child.data = data;
        if (data) {
          line.push(child);
          atLineStart = false;
          lastSpace = data.endsWith(' ');
        }
      } else if (child.nodeType === ELEMENT) {
        const tag = child.localName;
        if (isBlock(child)) {
          endBlock();
          if (tag !== 'pre' && tag !== 'hr') normalizeWhitespace(child);
        } else if (tag === 'br') {
          trimEnd(false);
          line.push(child);
          atLineStart = true;
          lastSpace = false;
        } else if (tag === 'img') {
          line.push(child);
          atLineStart = false;
          lastSpace = false;
        } else visit(child);
      }
    }
  };
  visit(container);
  endBlock();
}

function attrsKey(el) {
  return [...el.attributes].map(a => `${a.name}=${a.value}`).join(' ');
}

function mergeAdjacentInline(body) {
  for (const el of [...body.querySelectorAll(INLINE_SEL)]) {
    if (!el.isConnected) continue;
    for (;;) {
      let next = el.nextSibling;
      const gap = next?.nodeType === TEXT && !next.data.trim() ? next : null;
      if (gap) next = gap.nextSibling;
      if (!(next?.nodeType === ELEMENT && next.localName === el.localName && attrsKey(next) === attrsKey(el))) break;
      if (gap) el.append(gap);
      el.append(...next.childNodes);
      next.remove();
    }
  }
}

// ===================== Serializer =====================

function escapeText(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(value) {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

function openTag(el) {
  return `<${el.localName}${[...el.attributes].map(a => ` ${a.name}="${escapeAttr(a.value)}"`).join('')}>`;
}

function inlineHtml(nodes) {
  let html = '';
  for (const node of nodes) {
    if (node.nodeType === TEXT) html += escapeText(node.data.replace(/\u00a0/g, ' '));
    else if (node.nodeType === ELEMENT) {
      if (node.localName === 'br' || node.localName === 'img') html += openTag(node);
      else html += `${openTag(node)}${inlineHtml(node.childNodes)}</${node.localName}>`;
    }
  }
  return html;
}

function serialize(body) {
  const lines = [];
  for (const child of body.childNodes) serializeNode(child, 0, lines);
  return lines.length ? `${lines.join('\n')}\n` : '';
}

function serializeNode(node, depth, lines) {
  const pad = '  '.repeat(depth);
  if (!isBlock(node)) {
    const html = inlineHtml([node]).trim();
    if (html) lines.push(pad + html);
    return;
  }
  const tag = node.localName;
  if (tag === 'hr') { lines.push(`${pad}<hr>`); return; }
  if (tag === 'pre') {
    const text = escapeText(node.textContent.replace(/\u00a0/g, ' '));
    lines.push(`${pad}<pre>${text.startsWith('\n') ? '\n' : ''}${text}</pre>`);
    return;
  }
  const kids = [...node.childNodes];
  if (!kids.some(isBlock)) {
    lines.push(`${pad}${openTag(node)}${inlineHtml(kids)}</${tag}>`);
    return;
  }
  // Mixed content: a leading inline run shares the opening line; blocks are indented below it.
  const items = [];
  let run = [];
  const flush = () => {
    const html = inlineHtml(run).trim();
    if (html) items.push(html);
    run = [];
  };
  for (const kid of kids) {
    if (isBlock(kid)) { flush(); items.push(kid); }
    else run.push(kid);
  }
  flush();
  const first = typeof items[0] === 'string' ? items.shift() : '';
  lines.push(`${pad}${openTag(node)}${first}`);
  for (const item of items) {
    if (typeof item === 'string') lines.push(`${pad}  ${item}`);
    else serializeNode(item, depth + 1, lines);
  }
  lines.push(`${pad}</${tag}>`);
}
