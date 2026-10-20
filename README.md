# HTML Cleaner

Paste content copied from a web page, Word, or Google Docs and get clean, basic HTML that keeps the original's structure (headings, lists, tables, links, bold/italic…) with no CSS, no `style` attributes, and only the minimum of other attributes.

It's a single static page with no build step and no server-side logic, so it can be deployed to any static host.

> The full behavioral contract lives in [`SPEC.md`](SPEC.md). This README covers how to use, run, and work on the project.

## Using it

1. Copy something from Word, Google Docs, or a web page.
2. Paste it into the white **Paste** strip at the top. It's cleaned straight away; there's no button to press. You can also drag and drop content onto the strip.
3. Check the output, edit it if needed, then click **Copy HTML**.

| Control | What it does |
|---|---|
| **Source / WYSIWYG** | Switches the output between editable HTML source (the default) and a rich-text editor. Both views stay in sync. |
| **Re-clean** | Runs the cleaner again on the current Source contents. Hand edits are never re-cleaned unless you click this. |
| **Clear** | Empties the paste strip and the output. |
| **Debug** | Opens a drawer showing the raw clipboard HTML. Use it to capture new test fixtures (see below). |
| **Copy HTML** | Copies exactly what's in the output and briefly shows "Copied!". |
| **Drag bar** | Resizes the paste strip. Double-click to reset it, or focus it and use the arrow keys. The height is saved in `localStorage`. |

The **status bar** at the bottom shows the detected source (Word, Google Docs, Web page, Plain text…), the output's line and character counts, any warnings, and the current view.

### Warnings you may see

- **Emptied image sources.** Images embedded as `data:`, `file:`, or `blob:` URLs can't survive outside the original document, so the `<img>` is kept with `src=""`. In WYSIWYG view, use the image button to fill in a real URL.
- **"Switching views removed: …"** You hand-edited the Source view with tags or attributes the editor doesn't allow (for example `<div>` or `style` attributes), and switching to WYSIWYG dropped them.
- **"WYSIWYG editor could not load"** The editor (Tiptap) is loaded from `esm.sh` and needs internet access. Everything else, including cleaning, Source view, and Copy HTML, works offline.

## Running locally

Requires Node.js 22 or later.

```sh
npm install
npm start      # serves the app at http://localhost:8080 with caching disabled
npm test       # runs the cleaner's regression tests
```

The page loads `cleaner.js` as an ES module, so open it through the server rather than with `file://`.

## Project layout

```
index.html              UI: layout, styles, paste handling, Tiptap editor, toolbar, status bar
cleaner.js              Cleaning logic as a plain ES module, with no UI coupling
tests/cleaner.test.js   Fixture-driven and rule-focused tests (node:test)
tests/fixtures/         Paired raw-input / expected-output files
SPEC.md                 The specification: cleaning rules, UI, and testing approach
project-description.md  The original one-paragraph project brief
docs/agents/            Conventions for the coding-agent skills in .agents/skills/
```

`.scratch/` is git-ignored and holds local notes, issues, and screenshots.

## How the cleaner works

`cleaner.js` parses input with the global `DOMParser` (native in the browser, provided by jsdom in tests), rewrites the DOM in a fixed series of passes, and then serializes it with its own pretty-printer.

### Public API

| Export | Purpose |
|---|---|
| `cleanHtml(html)` | Cleans clipboard HTML. Returns `{ html, emptiedImages }`. |
| `cleanText(text)` | Cleans a plain-text paste. Blank lines separate paragraphs and single newlines become `<br>`. Returns the same shape as `cleanHtml`. |
| `clean(html)` | Shorthand for `cleanHtml(html).html`. |
| `detectSource(html)` | Returns `'Word'`, `'Google Docs'`, `'Excel'`, `'Google Sheets'`, or `'Web page'`. |
| `listDisallowed(html)` | Lists what the allowed set would drop, e.g. `['<div>', 'style attributes']`. The UI uses this for the view-switch warning. |

### Pipeline

`cleanHtml` runs these passes in order:

1. **Source detection.** Word and Google Docs input get extra source-specific passes.
2. **Word VML shapes** (Word only). WordArt text is pulled out of conditional comments, and shape fallback images are dropped.
3. **Junk removal.** `script`, `style`, comments, and similar content are removed. `pre` is flattened to plain text.
4. **Word clean-up** (Word only). This pass handles comments, tracked changes (insertions kept, deletions removed), footnotes and endnotes, text boxes, table-of-contents page numbers, real lists rebuilt from `mso-list` paragraphs and literal bullets, and `MsoHeadingN`/`MsoTitle` headings.
5. **Google Docs titles** become `h1`, and footnote links become `<sup>` marks.
6. **Page breaks and TOC dot leaders** are removed.
7. **Formatting.** `b`→`strong`, `i`→`em`, `del`/`strike`→`s`, and so on. Bold and italic inline styles become tags before styles are stripped.
8. **Namespaced tags** (`o:p`, `w:*`…) are removed. Unsafe or anchor-only links are unwrapped, and image sources are fixed.
9. **Structure.** Disallowed tags are unwrapped, `div`-like blocks become `p`, inline elements that wrap blocks are pushed down into each block, and loose inline content gets wrapped.
10. **Attribute whitelist.** Only the attributes listed in `ALLOWED_ATTRS` survive.
11. **Tidy, repeated until stable (at most 8 rounds).** This pass normalizes whitespace, removes empty inline elements and blocks, merges adjacent identical inline tags, and drops bold that covers a whole heading.
12. **Serialize.** Each block goes on its own line, nesting is indented by 2 spaces, and Unicode is written as real characters rather than entities.

The cleaner is **idempotent**: `clean(clean(x)) === clean(x)`. This matters because the WYSIWYG editor's output and the **Re-clean** button both pass already-clean HTML back through it, and it must not drift. The tests enforce this for every fixture.

### How the UI uses it

- A paste calls `cleanHtml` when the clipboard has `text/html`, or `cleanText` when it only has plain text.
- The paste strip shows a sanitized rendering of the original. Scripts, forms, embeds, and positioning are stripped, and attributes are whitelisted.
- Every WYSIWYG edit goes `editor.getHTML()` → `clean()` → Source view. The editor's schema is limited to the allowed tags, and links are configured to emit no `rel` or `class`.

## Testing

Only `cleaner.js` is tested automatically. The UI (pasting, the editor, view sync, copying) is checked by hand in the preview server.

Tests use Node's built-in `node:test` and `node:assert`, with jsdom supplying `DOMParser`. There are two kinds:

- **Fixtures.** Files in `tests/fixtures/` are discovered automatically, one test per input:
  - `<name>.input.html` holds raw clipboard HTML and goes through `cleanHtml`.
  - `<name>.input.txt` holds a plain-text paste and goes through `cleanText`.
  - `<name>.expected.html` holds the exact expected output, compared as a string including line breaks and indentation. Each fixture also checks `clean(expected) === expected`.
- **Focused rule tests** in `tests/cleaner.test.js`. These are small inline cases, one per cleaning rule.

### Adding a fixture from a real paste

1. Run `npm start`, turn on **Debug**, and paste the content.
2. Copy the raw HTML from the debug drawer into `tests/fixtures/<name>.input.html`.
3. Check the cleaned output carefully, then save it as `tests/fixtures/<name>.expected.html`.
4. Run `npm test`.

The current fixtures (`word-mac`, `google-docs`, `web-article`, `web-table`, `plain-text`) were written by hand from each source's known output. Replace or add to them with real clipboard captures over time.

## Changing the rules

- Allowed tags and attributes are defined by `BLOCK_TAGS`, `INLINE_TAGS`, and `ALLOWED_ATTRS` at the top of `cleaner.js`. If you change them, update the Tiptap extensions in `index.html` and the tables in `SPEC.md` to match, or the view-switch warning and the editor will disagree with the cleaner.
- Add a focused test for every new rule, and check that the idempotence assertion still passes.
