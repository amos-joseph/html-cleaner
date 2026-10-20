import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

globalThis.DOMParser = new JSDOM().window.DOMParser;

const { clean, cleanHtml, cleanText, detectSource, listDisallowed } = await import('../cleaner.js');

// ---------- Fixtures: <name>.input.html | <name>.input.txt → <name>.expected.html ----------

const dir = new URL('./fixtures/', import.meta.url);
const read = name => readFileSync(new URL(name, dir), 'utf8');

for (const file of readdirSync(dir).sort()) {
  const match = file.match(/^(.+)\.input\.(html|txt)$/);
  if (!match) continue;
  const [, name, kind] = match;
  test(`fixture: ${name}`, () => {
    const expected = read(`${name}.expected.html`);
    const actual = kind === 'html' ? cleanHtml(read(file)).html : cleanText(read(file)).html;
    assert.equal(actual, expected);
    assert.equal(clean(expected), expected, 'cleaning the expected output must not change it');
  });
}

// ---------- Focused rules ----------

test('b/i become strong/em; inline bold and italic styles are converted', () => {
  assert.equal(
    clean('<p><b>a</b> <i>b</i> <span style="font-weight:700">c</span> <span style="font-style: italic">d</span></p>'),
    '<p><strong>a</strong> <em>b</em> <strong>c</strong> <em>d</em></p>\n',
  );
});

test('b with font-weight normal (Google Docs wrapper) is not bold', () => {
  assert.equal(clean('<b style="font-weight:normal;" id="docs-internal-guid-1"><p>x</p></b>'), '<p>x</p>\n');
});

test('disallowed tags are unwrapped and divs become paragraphs', () => {
  assert.equal(
    clean('<div>one <span class="x">two</span></div><section><div>three</div></section><font color="red">four</font>'),
    '<p>one two</p>\n<p>three</p>\n<p>four</p>\n',
  );
});

test('only allowed attributes survive', () => {
  assert.equal(
    clean('<p id="a" class="b" style="color:red" lang="en" data-x="1" onclick="x()">t <a href="https://e.com" rel="nofollow" target="_blank" title="T" class="l">l</a></p>'),
    '<p>t <a href="https://e.com" target="_blank" title="T">l</a></p>\n',
  );
});

test('javascript:, anchor-only and empty-href links are unwrapped', () => {
  assert.equal(
    clean('<p><a href="javascript:alert(1)">a</a> <a name="x">b</a> <a href="">c</a> <a href=" JavaScript:void(0)">d</a></p>'),
    '<p>a b c d</p>\n',
  );
});

test('images: http kept, data/file/blob emptied and counted, alt added', () => {
  const { html, emptiedImages } = cleanHtml(
    '<p><img src="https://e.com/a.png" width="10"><img src="data:image/png;base64,AAA" alt="x"><img src="file:///C:/a.png"><img src="blob:https://e.com/1"></p>',
  );
  assert.equal(html, '<p><img src="https://e.com/a.png" alt=""><img src="" alt="x"><img src="" alt=""><img src="" alt=""></p>\n');
  assert.equal(emptiedImages, 3);
});

test('table cell spans only when > 1; ol start/type only when not default', () => {
  assert.equal(
    clean('<table><tr><td colspan="1" rowspan="2">a</td><td colspan="2">b</td></tr></table><ol start="1" type="1"><li>x</li></ol><ol start="3" type="a"><li>y</li></ol>'),
    [
      '<table>',
      '  <tbody>',
      '    <tr>',
      '      <td rowspan="2">a</td>',
      '      <td colspan="2">b</td>',
      '    </tr>',
      '  </tbody>',
      '</table>',
      '<ol>',
      '  <li>x</li>',
      '</ol>',
      '<ol start="3" type="a">',
      '  <li>y</li>',
      '</ol>',
      '',
    ].join('\n'),
  );
});

test('empty paragraphs and empty inline elements are removed', () => {
  assert.equal(clean('<p></p><p>&nbsp;</p><p><br></p><p>a<strong></strong><a href="https://e.com"></a></p>'), '<p>a</p>\n');
});

test('adjacent identical inline tags are merged', () => {
  assert.equal(clean('<p><strong>a</strong><strong>b</strong> <em>c</em> <em>d</em></p>'), '<p><strong>ab</strong> <em>c d</em></p>\n');
});

test('whitespace: nbsp becomes space, runs collapse, blocks are trimmed, trailing br removed', () => {
  assert.equal(clean('<p>  a&nbsp;&nbsp; b <br> c<br><br></p>'), '<p>a b<br>c</p>\n');
});

test('unicode is output as characters, not entities', () => {
  assert.equal(clean('<p>&ldquo;Hi&rdquo; &mdash; caf&eacute; &amp; &lt;tag&gt;</p>'), '<p>“Hi” — café &amp; &lt;tag&gt;</p>\n');
});

test('headings are never inferred from size or bold', () => {
  assert.equal(clean('<p style="font-size:28pt"><b>Big</b></p>'), '<p><strong>Big</strong></p>\n');
});

test('script, style and comments are removed', () => {
  assert.equal(clean('<style>p{}</style><script>x()</script><!-- c --><p>a</p>'), '<p>a</p>\n');
});

test('pre keeps its text and line breaks', () => {
  assert.equal(clean('<pre><code><span class="k">if</span> (a) {\n  b();\n}</code></pre>'), '<pre>if (a) {\n  b();\n}</pre>\n');
});

test('plain text: blank lines split paragraphs, single newlines become br, nothing inferred', () => {
  assert.equal(
    cleanText('First line\nsecond line\n\n\n- not a list\nhttps://example.com <b>\n').html,
    '<p>First line<br>second line</p>\n<p>- not a list<br>https://example.com &lt;b&gt;</p>\n',
  );
});

test('empty input gives empty output', () => {
  assert.equal(clean(''), '');
  assert.equal(cleanText('  \n\n ').html, '');
});

test('detectSource', () => {
  assert.equal(detectSource('<html xmlns:w="urn:schemas-microsoft-com:office:word"><p class=MsoNormal>x</p>'), 'Word');
  assert.equal(detectSource('<b id="docs-internal-guid-abc"><p>x</p></b>'), 'Google Docs');
  assert.equal(detectSource('<p>x</p>'), 'Web page');
});

test('listDisallowed names tags and attributes the editor schema would drop', () => {
  assert.deepEqual(listDisallowed('<div style="x"><p class="a">t</p><span>s</span></div><td colspan="2">c</td>'), [
    '<div>',
    '<span>',
    'class attributes',
    'style attributes',
  ]);
  assert.deepEqual(listDisallowed('<p>fine <a href="https://e.com" title="t">ok</a></p>'), []);
});

test('Word WordArt text becomes a paragraph and its fallback image is dropped', () => {
  const html = `<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><p class=MsoNormal>Before<!--[if gte vml 1]><v:shape id="WordArt_x0020_1" type="#_x0000_t136"><v:textpath style='font-family:"Calibri"' string="Big &amp; Bold"/></v:shape><![endif]--><![if !vml]><img src="file:///x.png" v:shapes="WordArt_x0020_1"><![endif]> after</p></body></html>`;
  const { html: out, emptiedImages } = cleanHtml(html);
  assert.equal(out, '<p>Before</p>\n<p>Big &amp; Bold</p>\n<p>after</p>\n');
  assert.equal(emptiedImages, 0);
});

test('Word paragraphs starting with a literal bullet become a list', () => {
  assert.equal(
    clean('<p class=MsoNormal>•&nbsp;&nbsp;One</p><p class=MsoNormal>• Two</p><p class=MsoNormal>After</p>'),
    '<ul>\n  <li>One</li>\n  <li>Two</li>\n</ul>\n<p>After</p>\n',
  );
});

test('Word roman and upper-alpha numbering keep their type', () => {
  const item = (marker, text) => `<p class=MsoListParagraph style='mso-list:l0 level1 lfo1'><span style='mso-list:Ignore'>${marker}&nbsp;</span>${text}</p>`;
  assert.equal(clean(item('ii.', 'Two') + item('iii.', 'Three')), '<ol start="2" type="i">\n  <li>Two</li>\n  <li>Three</li>\n</ol>\n');
  assert.equal(clean(item('A)', 'First')), '<ol type="A">\n  <li>First</li>\n</ol>\n');
});

test('deletions are only removed for Word; elsewhere del is strikethrough', () => {
  assert.equal(clean('<p>a <del>b</del> <ins>c</ins></p>'), '<p>a <s>b</s> c</p>\n');
});

test('Google Docs footnote marks become sup and lose their links', () => {
  assert.equal(
    clean('<p>Text<sup><a href="#ftnt1" id="ftnt_ref1">[1]</a></sup></p><div><p><a href="#ftnt_ref1" id="ftnt1">[1]</a> Note.</p></div>'),
    '<p>Text<sup>1</sup></p>\n<p><sup>1</sup> Note.</p>\n',
  );
});

test('Word table-of-contents dot leaders and page numbers are stripped, other paragraphs are untouched', () => {
  assert.equal(
    clean('<p class=MsoToc1>Introduction ............ 3</p><p class=MsoNormal>Count 1, 2, 3, .... 100</p>'),
    '<p>Introduction</p>\n<p>Count 1, 2, 3, .... 100</p>\n',
  );
});

test('long runs of spaces or dots are handled in linear time', () => {
  const start = Date.now();
  clean(`<p class=MsoToc1>a${' '.repeat(40000)}x</p><p>${'. '.repeat(20000)}</p>`);
  assert.ok(Date.now() - start < 1000);
});

test('button and label text is kept; zero-width-only paragraphs and href="#" links are removed', () => {
  assert.equal(clean('<p>Click <button>here</button> <a href="#">now</a></p><p>\u200b</p>'), '<p>Click here now</p>\n');
});

test('an inline element wrapping blocks is pushed into each block', () => {
  assert.equal(
    clean('<a href="https://e.com"><h3>Title</h3><div>Teaser</div></a>'),
    '<h3><a href="https://e.com">Title</a></h3>\n<p><a href="https://e.com">Teaser</a></p>\n',
  );
});

test('bold that covers a whole heading is dropped', () => {
  assert.equal(clean('<h2><b>All bold</b></h2><h2>Part <b>bold</b></h2>'), '<h2>All bold</h2>\n<h2>Part <strong>bold</strong></h2>\n');
});

test('editor (Tiptap-style) output is normalized to the same shape', () => {
  const tiptap = '<h1>T</h1><ul><li><p>a</p><ul><li><p>b</p></li></ul></li></ul><table style="min-width: 50px"><colgroup><col style="min-width: 25px"><col></colgroup><tbody><tr><th colspan="1" rowspan="1"><p>h</p></th><td colspan="1" rowspan="1" data-colwidth="100"><p>c</p></td></tr></tbody></table><pre><code class="language-js">x = 1</code></pre><p><a target="_blank" rel="noopener noreferrer nofollow" href="https://e.com">l</a><img src="" alt="x"></p><p></p>';
  assert.equal(
    clean(tiptap),
    [
      '<h1>T</h1>',
      '<ul>',
      '  <li>a',
      '    <ul>',
      '      <li>b</li>',
      '    </ul>',
      '  </li>',
      '</ul>',
      '<table>',
      '  <tbody>',
      '    <tr>',
      '      <th>h</th>',
      '      <td>c</td>',
      '    </tr>',
      '  </tbody>',
      '</table>',
      '<pre>x = 1</pre>',
      '<p><a href="https://e.com" target="_blank">l</a><img src="" alt="x"></p>',
      '',
    ].join('\n'),
  );
});

test('multi-paragraph cells and items keep their paragraphs', () => {
  const out = clean('<table><tr><td><p>a</p><p>b</p></td></tr></table>');
  assert.equal(out, '<table>\n  <tbody>\n    <tr>\n      <td>\n        <p>a</p>\n        <p>b</p>\n      </td>\n    </tr>\n  </tbody>\n</table>\n');
  assert.equal(clean(out), out);
});
