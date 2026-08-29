// render-doc.js — the .md IS the source, for supporting documents (case studies,
// one-pagers) that aren't résumés or letters.
//
//   node render-doc.js "<path/to/Doc.md>" [--out X.docx]
//
// Contract:
//   # Title                    -> document title
//   *subtitle line*            -> subtitle (italic, directly under the title)
//   a contact line in the head -> byline
//   ## SECTION                 -> section heading, uppercased, with rule
//   - item                     -> bullet
//   anything else              -> body paragraph
//   a bare --- after content   -> end of document

const { Document, Packer, Paragraph, TextRun, BorderStyle } = require('docx');
const fs = require('fs');
const path = require('path');

function runs(text, base = {}) {
  return text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g).filter(Boolean).map(p => {
    if (p.startsWith('**') && p.endsWith('**')) return new TextRun({ text: p.slice(2, -2), bold: true, ...base });
    if (p.startsWith('*')  && p.endsWith('*'))  return new TextRun({ text: p.slice(1, -1), italics: true, ...base });
    return new TextRun({ text: p, ...base });
  });
}
const ats = t => t.replace(/\s*[→⟶]\s*/g, ' to ').replace(/\s*↔\s*/g, ' and ');

// A bare `---` ends the document ONLY when nothing but blank lines and an italic
// footer note follow it. A `---` with real content after it is a section divider
// (case studies and master résumés use one under the header block).
function bodyEnd(lines) {
  for (let i = lines.length - 1; i >= 0; i--) {
    if (!/^---+$/.test(lines[i].trim())) continue;
    const after = lines.slice(i + 1).map(l => l.trim()).filter(Boolean);
    if (after.every(l => /^\*[^*][^*]*\*$/.test(l))) return i;
    break;
  }
  return lines.length;
}

const widen = t => t.split(/(\*\*[^*]+\*\*)/g).map(p => (p.startsWith('**') ? p : p.replace(/ · /g, '  ·  '))).join('');

const args = process.argv.slice(2);
const mdPath = args.find(a => !a.startsWith('--'));
if (!mdPath) { console.error('usage: render-doc.js <Doc.md> [--out X.docx]'); process.exit(2); }
const flag = n => { const i = args.indexOf('--' + n); return i >= 0 ? args[i + 1] : null; };
const out = flag('out') || path.join(path.dirname(mdPath), path.basename(mdPath, '.md') + '.docx');

let src = fs.readFileSync(mdPath, 'utf8');
if (src.startsWith('---')) src = src.split(/^---\s*$/m).slice(2).join('---');

const lines = src.split('\n');
const starts = lines.map((l, i) => (/^#\s+/.test(l.trim()) ? i : -1)).filter(i => i >= 0);
if (starts.length > 1) lines.splice(0, starts[starts.length - 1]);

const nodes = [];
let inHead = false, sawTitle = false;
for (const raw of lines.slice(0, bodyEnd(lines))) {
  const l = raw.trim();
  if (!l) continue;
  if (/^---+$/.test(l)) continue;
  if (l.startsWith('>')) continue;
  if (l.startsWith('# ')) { nodes.push({ t: 'title', v: l.slice(2).trim() }); inHead = true; sawTitle = true; continue; }
  if (/^#{2,6}\s/.test(l)) { inHead = false; nodes.push({ t: 'heading', v: l.replace(/^#+\s*/, '').trim() }); continue; }
  if (inHead) {
    if (/^\*[^*][^*]*\*$/.test(l)) { nodes.push({ t: 'subtitle', v: l.slice(1, -1).trim() }); continue; }
    // a byline is a contact line; anything else under the title is real prose
    if (/@|\d{3}[-.]\d{3}|linkedin\.com/i.test(l)) { nodes.push({ t: 'byline', v: l }); continue; }
    inHead = false;
    nodes.push({ t: 'body', v: l });
    continue;
  }
  if (/^[-*+]\s+/.test(l)) { nodes.push({ t: 'bullet', v: l.replace(/^[-*+]\s+/, '') }); continue; }
  nodes.push({ t: 'body', v: l });
}
if (!sawTitle) { console.error('ERROR: no "# Title" heading found in ' + mdPath); process.exit(1); }

// headings render uppercase; strip inline emphasis and any trailing parenthetical aside
const headingText = v => ats(v).replace(/[*_`]/g, '').replace(/\s*\([^)]*\)\s*$/, '').toUpperCase();

const P = {
  title:    v => new Paragraph({ spacing: { after: 8 },  children: [new TextRun({ text: ats(v), bold: true, size: 30 })] }),
  subtitle: v => new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: ats(v), italics: true, size: 19, color: '333333' })] }),
  byline:   v => new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: widen(ats(v)), size: 18, color: '444444' })] }),
  heading:  v => new Paragraph({ spacing: { before: 200, after: 60 },
                                 border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '888888', space: 1 } },
                                 children: [new TextRun({ text: headingText(v), bold: true, size: 22, color: '1a1a1a' })] }),
  body:     v => new Paragraph({ spacing: { after: 90, line: 264 }, children: runs(widen(ats(v)), { size: 20 }) }),
  bullet:   v => new Paragraph({ bullet: { level: 0 }, spacing: { after: 66, line: 264 }, children: runs(ats(v), { size: 20 }) }),
};

const doc = new Document({
  styles: { default: { document: { run: { font: 'Calibri', size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 720, bottom: 720, left: 936, right: 936 } } },
    children: nodes.map(n => P[n.t](n.v)),
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(out, b);
  const c = nodes.reduce((a, n) => ((a[n.t] = (a[n.t] || 0) + 1), a), {});
  console.log(`WROTE ${path.basename(out)} (${b.length}b) ` + Object.entries(c).map(([k, v]) => `${k}:${v}`).join(' '));
});
