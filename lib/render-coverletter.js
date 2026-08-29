// render-coverletter.js — the .md IS the source, same contract as render-resume.js.
//
//   node render-coverletter.js "<path/to/Cover Letter - Foo.md>" [--out X.docx] [--date "August 5, 2026"]
//
// The letterhead (name, contact, date) is synthesized: the .md holds the letter itself,
// starting at the greeting. The date comes from the frontmatter `timestamp:` unless
// --date overrides it, so the source governs that too.
//
// Not rendered: frontmatter, the `# Title` line, `>` editorial banners, and the
// `---` rule that separates them from the letter.

const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs');
const path = require('path');

const CFG = (() => {
  const i = process.argv.indexOf('--config');
  const f = i >= 0 ? process.argv[i + 1] : 'docgate.config.json';
  return fs.existsSync(f) ? (JSON.parse(fs.readFileSync(f, 'utf8')).author || {}) : {};
})();
const NAME = (CFG.name || 'YOUR NAME').toUpperCase();
const DEFAULT_LOCATION = 'Portland, OR (Remote)';
const contactLine = loc => [loc, ...(CFG.contact || [])].filter(Boolean).join('  ·  ');
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

function runs(text, base = {}) {
  return text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g).filter(Boolean).map(p => {
    if (p.startsWith('**') && p.endsWith('**')) return new TextRun({ text: p.slice(2, -2), bold: true, ...base });
    if (p.startsWith('*')  && p.endsWith('*'))  return new TextRun({ text: p.slice(1, -1), italics: true, ...base });
    return new TextRun({ text: p, ...base });
  });
}
const ats = t => t.replace(/\s*[→⟶]\s*/g, ' to ').replace(/\s*↔\s*/g, ' and ');

function parse(md) {
  const fm = md.match(/^---\r?\n([\s\S]*?)\r?\n---/);
  const ts = fm && fm[1].match(/timestamp:\s*"?([0-9]{4})-([0-9]{2})-([0-9]{2})"?/);
  const locM = fm && fm[1].match(/^location:\s*"?(.+?)"?\s*$/m);
  const location = locM ? locM[1] : DEFAULT_LOCATION;
  const date = ts ? `${MONTHS[+ts[2] - 1]} ${+ts[3]}, ${ts[1]}` : null;

  const body = fm ? md.slice(md.indexOf('---', 3) + 3) : md;
  const raw = body.split('\n').map(l => l.trim());

  // A letter begins at its salutation. Everything above it — doc title, word-count
  // notes, editorial banners, the rule that separates them — is preamble, whatever
  // form it takes. Anchoring here is what makes leading notes structurally harmless.
  const GREETING = /^(Dear|Hello|Hi|Greetings|To the|To whom)\b.*[,:]\s*$/i;
  const start = raw.findIndex(l => GREETING.test(l));
  if (start < 0) return { date, location, lines: [], noGreeting: true };

  const lines = [];
  for (const l of raw.slice(start)) {
    if (!l) continue;
    if (/^---+$/.test(l)) break;                             // trailing notes never render
    if (l.startsWith('>') || /^#{1,6}\s/.test(l)) continue;
    if (/delete before sending/i.test(l)) continue;          // belt and braces
    lines.push(l);
  }
  return { date, location, lines };
}

const args = process.argv.slice(2);
const mdPath = args.find(a => !a.startsWith('--'));
if (!mdPath) { console.error('usage: render-coverletter.js <Cover Letter.md> [--out X.docx] [--date "Month D, YYYY"]'); process.exit(2); }
const flag = n => { const i = args.indexOf('--' + n); return i >= 0 ? args[i + 1] : null; };
const out = flag('out') || path.join(path.dirname(mdPath), path.basename(mdPath, '.md') + '.docx');

const { date, location, lines, noGreeting } = parse(fs.readFileSync(mdPath, 'utf8'));
if (noGreeting) { console.error('ERROR: no salutation (Dear/Hello/Hi...) found in ' + mdPath + ' — cannot locate where the letter begins'); process.exit(1); }
const when = flag('date') || date;
if (!when) { console.error('ERROR: no frontmatter timestamp and no --date given'); process.exit(1); }
if (!lines.length) { console.error('ERROR: no letter body found in ' + mdPath); process.exit(1); }

// signature block = a trailing line matching the author name, plus whatever follows it
const sigAt = lines.findIndex(l => /^${CFG.name}$/i.test(l));
const bodyLines = sigAt > 0 ? lines.slice(0, sigAt - 1) : lines;
const closing  = sigAt > 0 ? lines[sigAt - 1] : null;
const sigRest  = sigAt > 0 ? lines.slice(sigAt + 1) : [];

const children = [
  new Paragraph({ spacing: { after: 0 },   children: [new TextRun({ text: NAME, bold: true, size: 32 })] }),
  new Paragraph({ spacing: { after: 320 }, children: [new TextRun({ text: contactLine(location), size: 18, color: '444444' })] }),
  new Paragraph({ spacing: { after: 260 }, children: [new TextRun({ text: when, size: 20, color: '333333' })] }),
  ...bodyLines.map(l => new Paragraph({ spacing: { after: 200, line: 276 }, children: runs(ats(l), { size: 21 }) })),
];
if (closing) {
  children.push(new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun({ text: closing, size: 21 })] }));
  children.push(new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun({ text: '${CFG.name}', size: 21, bold: true })] }));
  for (const r of sigRest)
    children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: r.replace(/ · /g, '  ·  '), size: 19, color: '444444' })] }));
}

const doc = new Document({
  styles: { default: { document: { run: { font: 'Calibri', size: 21 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 900, bottom: 900, left: 1080, right: 1080 } } },
    children,
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(out, b);
  console.log(`WROTE ${path.basename(out)} (${b.length}b, ${when}, ${bodyLines.length} paras)`);
});
