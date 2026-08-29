// render-resume.js — the .md IS the source. This parses it; it does not restate it.
//
//   node render-resume.js "<path/to/Resume - Foo.md>" [--out X.docx] [--density tight|loose]
//
// Replaces the per-package hardcoded build-*-resume.js scripts. Because the text
// lives in exactly one place, a render can no longer diverge from the markdown.
//
// Markdown contract (see Rendering-Notes.md):
//   # NAME                     -> name line
//   **tagline**   (in header)  -> tagline, rendered unbold
//   next header line           -> contact line
//   ## SECTION                 -> section heading w/ rule
//   ### Company · Dates        -> company line
//   *italic line*              -> italic subtitle
//   **Role — Team · Dates**    -> role line   (whole line bold)
//   - item                     -> bullet
//   anything else              -> body paragraph
//   a bare --- after content   -> end of document (footer note is not rendered)

const { Document, Packer, Paragraph, TextRun, BorderStyle } = require('docx');
const fs = require('fs');
const path = require('path');

const DENSITY = {
  tight: { heading: { before: 104, after: 34 }, role: { before: 54, after: 6 }, company: { before: 30, after: 6 },
           italic: { after: 20 }, body: { after: 28 }, bullet: { after: 16 } },
  loose: { heading: { before: 130, after: 42 }, role: { before: 68, after: 8 }, company: { before: 68, after: 8 },
           italic: { after: 26 }, body: { after: 38 }, bullet: { after: 24 } },
};

// Widen " · " to "  ·  " only OUTSIDE bold spans. Inside a bold run the separator
// is joining parts of one title (Project Name · 2025) and stays narrow; outside it
// is separating list items (skills, contact fields) and gets air.
function widen(text) {
  return text.split(/(\*\*[^*]+\*\*)/g)
    .map(p => (p.startsWith('**') ? p : p.replace(/ · /g, '  ·  ')))
    .join('');
}

// Handles **bold** and *italic*; order matters so ** is consumed before *.
function runs(text, base = {}) {
  return text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g).filter(Boolean).map(p => {
    if (p.startsWith('**') && p.endsWith('**')) return new TextRun({ text: p.slice(2, -2), bold: true, ...base });
    if (p.startsWith('*')  && p.endsWith('*'))  return new TextRun({ text: p.slice(1, -1), italics: true, ...base });
    return new TextRun({ text: p, ...base });
  });
}

const isFullBold   = l => /^\*\*.+\*\*$/.test(l) && !l.slice(2, -2).includes('**');
const isFullItalic = l => /^\*[^*][^*]*\*$/.test(l) && !l.startsWith('**');

// ATS safety: arrow glyphs must never reach the PDF (they garble in parsers and
// screen readers). The .md may keep them for readability; the render spells them out.
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

function parse(md) {
  let src = md.replace(/^﻿/, '');
  if (src.startsWith('---')) src = src.split(/^---\s*$/m).slice(2).join('---'); // drop frontmatter

  // Some sources (baseline, master variants) carry a document title and an editorial
  // note above the résumé proper. The résumé begins at the LAST top-level heading;
  // everything before it is preamble and is not part of the document.
  const lines = src.split('\n');
  const starts = lines.map((l, i) => (/^#\s+/.test(l.trim()) ? i : -1)).filter(i => i >= 0);
  if (starts.length > 1) lines.splice(0, starts[starts.length - 1]);

  const out = [];
  let inHeader = false, sawName = false;

  for (const raw of lines.slice(0, bodyEnd(lines))) {
    const line = raw.trim();
    if (!line) continue;
    if (/^---+$/.test(line)) continue;
    if (/^>/.test(line)) continue;                             // editorial callouts never render

    if (line.startsWith('# ')) { out.push({ t: 'name', v: line.slice(2).trim() }); inHeader = true; sawName = true; continue; }
    if (line.startsWith('## '))  { inHeader = false; out.push({ t: 'heading', v: line.slice(3).trim() }); continue; }
    if (line.startsWith('### ')) { inHeader = false; out.push({ t: 'company', v: line.slice(4).trim() }); continue; }

    if (inHeader) {
      if (isFullBold(line))        out.push({ t: 'tagline',  v: line.slice(2, -2).trim() });
      else if (isFullItalic(line)) out.push({ t: 'subtitle', v: line.slice(1, -1).trim() });
      else                         out.push({ t: 'contact',  v: line });
      continue;
    }
    if (/^[-*+]\s+/.test(line)) { out.push({ t: 'bullet', v: line.replace(/^[-*+]\s+/, '') }); continue; }
    if (isFullItalic(line))    { out.push({ t: 'italic', v: line.slice(1, -1).trim() }); continue; }
    if (isFullBold(line))      { out.push({ t: 'role',   v: line.slice(2, -2).trim() }); continue; }
    out.push({ t: 'body', v: line });
  }
  return out;
}

function build(rawNodes, density) {
  const S = DENSITY[density];
  const nodes = rawNodes.map(n => ({ t: n.t, v: ats(n.v) }));
  const P = {
    name:    v => new Paragraph({ spacing: { after: 0 },  children: [new TextRun({ text: v, bold: true, size: 32 })] }),
    tagline: v => new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: widen(v), size: 20, color: '333333' })] }),
    subtitle: v => new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: v, italics: true, size: 19, color: '333333' })] }),
    contact: v => new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: widen(v), size: 18, color: '444444' })] }),
    heading: v => new Paragraph({ spacing: S.heading, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '888888', space: 1 } },
                                  children: [new TextRun({ text: v, bold: true, size: 22, color: '1a1a1a' })] }),
    company: v => new Paragraph({ spacing: S.company, children: [new TextRun({ text: widen(v), bold: true, size: 21 })] }),
    role:    v => new Paragraph({ spacing: S.role,    children: [new TextRun({ text: widen(v), bold: true, size: 20 })] }),
    italic:  v => new Paragraph({ spacing: S.italic,  children: [new TextRun({ text: v, italics: true, size: 19, color: '333333' })] }),
    body:    v => new Paragraph({ spacing: S.body,    children: runs(widen(v), { size: 20 }) }),
    bullet:  v => new Paragraph({ bullet: { level: 0 }, spacing: S.bullet, children: runs(v, { size: 20 }) }),
  };
  return nodes.map(n => P[n.t](n.v));
}

const args = process.argv.slice(2);
const mdPath = args.find(a => !a.startsWith('--'));
if (!mdPath) { console.error('usage: render-resume.js <Resume.md> [--out X.docx] [--density tight|loose]'); process.exit(2); }
const flag = n => { const i = args.indexOf('--' + n); return i >= 0 ? args[i + 1] : null; };
const density = flag('density') || 'tight';
const out = flag('out') || path.join(path.dirname(mdPath), path.basename(mdPath, '.md') + '.docx');

const nodes = parse(fs.readFileSync(mdPath, 'utf8'));
if (!nodes.some(n => n.t === 'name')) { console.error('ERROR: no "# NAME" heading found in ' + mdPath); process.exit(1); }

const doc = new Document({
  styles: { default: { document: { run: { font: 'Calibri', size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 576, bottom: 576, left: 792, right: 792 } } },
    children: build(nodes, density),
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(out, b);
  const counts = nodes.reduce((a, n) => ((a[n.t] = (a[n.t] || 0) + 1), a), {});
  console.log(`WROTE ${path.basename(out)} (${b.length}b, ${density}) ` +
              Object.entries(counts).map(([k, v]) => `${k}:${v}`).join(' '));
});
