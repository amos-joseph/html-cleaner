# HTML Cleaner — Specification

Paste content copied from a web page, Word, or Google Docs and get clean, basic HTML that matches the original's general formatting — no CSS, no `style`, minimal attributes.

## Delivery & files

- Static page, no build step, no server-side logic. Deployable to any static host.
- Files:
  - `index.html` — UI
  - `cleaner.js` — cleaning logic as a plain ES module, no UI coupling. Parses HTML via the global `DOMParser` (native in the browser, provided by jsdom in tests).
  - `package.json` — `"type": "module"`, `engines: { "node": ">=22" }`, devDependencies `jsdom` and `http-server`
  - `tests/cleaner.test.js` + `tests/fixtures/` — Node-based regression tests
- **Preview:** `npm start` runs `http-server . -c-1 -p 8080` and serves the app at `http://localhost:8080`. Caching is disabled (`-c-1`) so edits always show up.
- Tiptap (WYSIWYG editor) loaded from a CDN (e.g. `esm.sh`) via an import map. The editor needs internet access.

## Cleaning rules

### Allowed tags

| Group | Tags |
|---|---|
| Structure | `p`, `h1`–`h6`, `br`, `hr`, `blockquote` |
| Inline | `strong`, `em`, `u`, `s`, `sub`, `sup`, `code` |
| Lists | `ul`, `ol`, `li` (nesting allowed) |
| Links | `a` |
| Images | `img` |
| Tables | `table`, `thead`, `tbody`, `tr`, `th`, `td` |
| Preformatted | `pre` |

- Any other tag is unwrapped: its text is kept and the tag is dropped.
- `div` becomes `p` where it acts as a paragraph.
- Always removed: `style`, `script`, comments, `o:p` and other namespaced (Word XML) tags.
- `b` becomes `strong` and `i` becomes `em`.
- Inline styles `font-weight: bold/600+/700` and `font-style: italic` become `strong`/`em` before styles are stripped.

### Allowed attributes

Everything not listed is stripped (`class`, `style`, `id`, `lang`, `dir`, `width`/`height`, `data-*`, `mso-*`, event handlers, `rel`, `border`, …).

| Tag | Attributes |
|---|---|
| `a` | `href`, `target`, `title` |
| `img` | `src`, `alt` (`alt=""` added if missing) |
| `td`, `th` | `colspan`, `rowspan` — only when > 1 |
| `ol` | `start` (only when ≠ 1), `type` (only when ≠ `1`) |

- Links with `javascript:` hrefs, anchor-only links (`<a name>`), and empty-href links are unwrapped.

### Images

- `http(s)://` sources: kept as is.
- `data:` (base64), `file:`, and `blob:` sources: the `img` tag is kept with `src=""`, and a UI warning reports how many images were affected so a real URL can be filled in.

### Lists

- Word's fake lists (`MsoListParagraph`, `mso-list` levels, literal bullet/number characters) are rebuilt into real nested `ul`/`ol`.
- Numbering and list type are preserved via `start`/`type` where needed.

### Headings

- Only real `h1`–`h6` tags and named heading styles (`MsoHeading1`…, `MsoTitle`, Google Docs headings) become headings.
- No inference from font size or bolding. A large bold line becomes `<p><strong>…</strong></p>`.

### Whitespace & empty elements

- Empty paragraphs (`<p></p>`, `<p>&nbsp;</p>`, `<p><br></p>`) are removed.
- `&nbsp;` becomes a regular space, and runs of spaces collapse to one.
- Empty inline elements (e.g. `<strong></strong>`, links with no text) are removed.
- Adjacent identical inline tags are merged (`<strong>a</strong><strong>b</strong>` → `<strong>ab</strong>`).
- Leading and trailing whitespace inside blocks is trimmed.
- `<br>` is kept inside blocks, and trailing `<br>` at the end of a block is removed.
- Smart quotes, dashes, and other Unicode are output as real characters, not entities.

### Word / Google Docs extras

- **Footnotes and endnotes:** reference marks become `<sup>1</sup>`, and the note text becomes plain paragraphs at the end. Internal anchor links are dropped.
- **Comments:** removed.
- **Tracked changes:** only the final text is kept (insertions kept, deletions removed).
- **Text boxes, shapes, and WordArt:** their text becomes plain paragraphs in place, and the shapes are dropped.
- **Table of contents:** entries become plain paragraphs, or links if the source has them. Page numbers and dot leaders are stripped where detectable.
- **Tables from Excel or Sheets:** handled like any other table, with formatting, widths, and number styles stripped.
- **Page and section breaks:** removed.

### Plain-text pastes (no `text/html` on the clipboard)

- Blank-line-separated chunks become `<p>`, and single line breaks within a chunk become `<br>`.
- Nothing else is inferred: no auto-linking and no list detection.

### Output layout

- Each block element (`p`, `h*`, `li`, `tr`, `td`, …) starts on its own line.
- Nested blocks are indented by 2 spaces.
- Inline content stays on the same line as its block, with no hard-wrapping.

## UI

### Layout

A full-window, dark app-style layout (like a code editor). The page does not scroll; only the panes inside it do. From top to bottom:

1. **App bar:** the app name and the Source/WYSIWYG toggle on the left. Re-clean, Clear, the Debug toggle, and **Copy HTML** (the primary, highlighted button) on the right.
2. **Paste strip:** a short paste area (120px by default) with a small vertical "Paste" label, rendered on a white background. A drag bar between it and the output resizes it (double-click resets, arrow keys work when focused); the chosen height is remembered in `localStorage`, and the output always keeps at least 120px.
3. **Output:** fills all remaining height. The WYSIWYG toolbar appears directly above it only in WYSIWYG view. Source view is dark and monospace; WYSIWYG view is rendered on white.
4. **Debug drawer:** a fixed-height panel below the output showing the raw clipboard HTML, visible only while Debug is on.
5. **Status bar:** a thin bar across the bottom showing the paste's detected source (Word, Google Docs, Web page, Plain text), the output's line and character counts, any warnings, and the current view.

Colors stay neutral greys. The status bar is the same grey as the app bar and never changes color; warnings are marked by a ⚠ and slightly brighter text, not a colored background. The only accent color is the Copy HTML button.

### Paste area

- An editable box that shows the pasted content rendered.
- Pasting triggers cleaning immediately, with no button.

### Output

- A toggle switches between **Source** (the default) and **WYSIWYG**. Both views are editable and stay in sync.
- **Source view:** monospace and editable.
  - Hand edits are never re-cleaned automatically.
  - **Copy HTML** copies exactly what is shown.
- **WYSIWYG view:** a Tiptap editor whose schema is limited to the allowed tags and attributes above.
  - The link extension is configured to emit no `rel` and to set `target` only when chosen.
  - Editor output is passed through `cleaner.js` as a safety net.
- **View-switch warning:** when switching Source → WYSIWYG would drop content Tiptap's schema doesn't allow, a warning lists what was removed (e.g. "Switching views removed: `<div>`, `style` attributes").

### Toolbar (WYSIWYG)

- **Text:** Bold, Italic, Underline, Strikethrough, Subscript, Superscript, Inline code
- **Blocks:** Paragraph/H1–H6 dropdown, Blockquote, Code block (`pre`), Horizontal rule
- **Lists:** Bulleted, Numbered, Indent, Outdent
- **Links:** Add/edit (URL, title, open in new tab → `target="_blank"`), Remove link
- **Images:** Insert/edit by URL and alt text (also used to fill in emptied `src` values)
- **Tables:** Insert table, add/delete row, add/delete column, merge/split cells, toggle header row
- **History:** Undo, Redo
- **Clear formatting:** removes inline marks from the selection

### Buttons & extras

- **Copy HTML:** shows a brief "Copied!" confirmation.
- **Clear:** resets the paste area and output.
- **Re-clean:** runs the cleaner on the current Source contents, on demand.
- **Warnings:** shown in the status bar (emptied image sources, view-switch losses).
- **Debug toggle:** opens the debug drawer with the raw clipboard HTML so new test fixtures can be captured.

### Responsiveness

- The output section is stacked below the paste area at all widths.

## Testing

- All automated testing runs in Node.js (≥ 22) using the built-in `node:test` and `node:assert`. `npm test` runs `node --test`.
- jsdom provides `DOMParser` to `cleaner.js` during tests.
- Scope: only `cleaner.js` is tested automatically. The UI (paste handling, Tiptap, view sync, Copy HTML) is checked manually in the preview server.
- **Fixtures:** paired files in `tests/fixtures/`, discovered automatically, with one test per input:
  - `<name>.input.html` — raw clipboard HTML, cleaned via the HTML path
  - `<name>.input.txt` — plain-text paste, cleaned via the plain-text path
  - `<name>.expected.html` — the exact expected output, compared as a string (line breaks and indentation included)
- **Idempotence:** every fixture also asserts that cleaning its expected output returns it unchanged (`clean(expected) === expected`), so Re-clean and the WYSIWYG safety net never cause drift.
- Starter fixtures, written from the known output of each source:
  - Word desktop (Mac)
  - Google Docs
  - a web article
  - a web page with a table
  - a plain-text paste
- New fixtures are added by copying raw HTML from the debug drawer into a new `*.input.html` file. These will be replaced or supplemented with real clipboard captures made using the debug toggle.
