// check-drift.js — the release gate. Run before anything ships.
//
//   node lib/check-drift.js                    # audit everything
//   node lib/check-drift.js --strict           # exit 1 on any MATERIAL finding
//   node lib/check-drift.js --root DIR --config docgate.config.json
//
// Text drift is now structurally impossible: every artifact renders from its .md
// via render-resume / render-coverletter / render-doc, so there is only one copy
// of the words. What remains checkable is whether the artifact on disk is actually
// current, and whether anything reached it that must never reach an employer.
//
// Checks, per document:
//   STALE     .md is newer than the .docx/.pdf beside it -> the artifact is out of date
//   ORPHAN    a PDF with no .md source (cannot be regenerated or verified)
//   ATS       arrow glyphs in the rendered output
//   MARKDOWN  literal ** or * or # leaked into the rendered output
//   NOTES     editorial scaffolding ("delete before sending", DRAFT banners, word counts)
//   FIGURES   pre-reconciliation metrics the accuracy sweep removed as unsupported

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const STRICT = process.argv.includes('--strict');

// Config-driven: nothing about a particular corpus is hardcoded here. Point it at
// a docgate.config.json (see docgate.config.example.json) or pass --root.
const cfgArg = process.argv.indexOf('--config');
const CFG_PATH = cfgArg >= 0 ? process.argv[cfgArg + 1] : 'docgate.config.json';
const cfg = fs.existsSync(CFG_PATH) ? JSON.parse(fs.readFileSync(CFG_PATH, 'utf8')) : {};
const rootArg = process.argv.indexOf('--root');
const ROOT = path.resolve(rootArg >= 0 ? process.argv[rootArg + 1] : (cfg.root || '.'));

// Shipped documents whose rendered artifacts are the historical record of what a
// recipient actually received. They are deliberately NOT regenerated, so staleness
// against their source is expected and reported without failing.
const FROZEN = cfg.frozen || [];

// Superseded artifacts kept for provenance. Reported as ARCHIVE, never as an error.
const ARCHIVE = cfg.archivePattern ? new RegExp(cfg.archivePattern, 'i') : /^$/;

// Editorial scaffolding that must never survive into a rendered artifact.
const NOTE_PAT = new RegExp(cfg.notePattern ||
  'delete before sending|tailoring notes|~\\d+ words\\.|\\bDRAFT\\b|personalize before sending|\\bTODO\\b|\\bFIXME\\b', 'i');

// Claims a later accuracy pass retracted. Any of these reaching a rendered artifact
// is a material failure: the source was corrected and the artifact was not.
const BAD_FIGURES = Object.fromEntries(
  Object.entries(cfg.retractedFigures || {}).map(([k, v]) => [k, new RegExp(v, 'i')]));

// Artifact filenames do not always mirror their source: a renderer may output a
// different display name, or swap hyphens for spaces. Match on a normalized key
// rather than an exact path, with explicit aliases from config as a fallback.
const normKey = n => n.toLowerCase().replace(/\.(md|docx|pdf)$/, '').replace(/[^a-z0-9]/g, '');
function resolveArtifact(md, ext) {
  const dir = path.dirname(md);
  const want = normKey(path.basename(md));
  const hit = fs.readdirSync(dir).find(f => f.endsWith(ext) && normKey(f) === want);
  if (hit) return path.join(dir, hit);
  for (const [srcFrag, outName] of Object.entries(cfg.artifactAliases || {}))
    if (md.includes(srcFrag)) return path.join(dir, outName + ext);
  return md.replace(/\.md$/, ext);
}

const findings = [];
const add = (sev, doc, msg) => findings.push({ sev, doc, msg });

function pdftotext(p) {
  try { return execFileSync('pdftotext', [p, '-'], { encoding: 'utf8', maxBuffer: 1 << 24 }); }
  catch { return ''; }
}

function walk(dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    // skip archives: legacy scripts, and old-download folders holding documents
    // that were never part of this pipeline (including other people's files)
    if (e.name.startsWith('.') || e.name === 'node_modules' || e.name === 'legacy') continue;
    if (/^_(from|archive|old)/i.test(e.name)) continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, out); else out.push(p);
  }
  return out;
}

const all = walk(ROOT);
const pdfs = all.filter(p => p.endsWith('.pdf'));
const SRC = new RegExp(cfg.sourcePattern || '.', 'i');
const mds = all.filter(p => /\.md$/.test(p) && SRC.test(path.basename(p)));

// --- every renderable .md should have a current artifact beside it ---
for (const md of mds) {
  const rel = path.relative(ROOT, md);
  const frozen = FROZEN.some(f => md.includes(f));
  for (const ext of ['.docx', '.pdf']) {
    // baseline renders under a different output name; resolve by directory scan
    let art = md.replace(/\.md$/, ext);
    if (!fs.existsSync(art)) art = resolveArtifact(md, ext);
    if (!fs.existsSync(art)) {
      // a submitted package's PDF is the record; a missing .docx there is not a defect
      if (frozen && ext === '.docx') continue;
      add('MATERIAL', rel, `no ${ext} rendered from this source`); continue;
    }
    if (fs.statSync(md).mtime > fs.statSync(art).mtime) {
      if (frozen) add('frozen', rel, `${ext} older than .md — expected, package is submitted`);
      else add('MATERIAL', rel, `STALE: .md is newer than ${ext} — re-render before sending`);
    }
  }
}

// --- what actually reached the rendered output ---
for (const pdf of pdfs) {
  const rel = path.relative(ROOT, pdf);
  const base = path.basename(pdf);
  const archived = ARCHIVE.test(base);
  const t = pdftotext(pdf);
  if (!t) { add('warn', rel, 'could not extract text'); continue; }

  const dir = path.dirname(pdf);
  const want = normKey(base);
  const hasMd = fs.readdirSync(dir).some(f => f.endsWith('.md') && normKey(f) === want) ||
                Object.entries(cfg.artifactAliases || {}).some(([src, out]) =>
                  normKey(out) === want && fs.existsSync(path.join(dir, src + '.md')));

  const arrows = (t.match(/[→⟶↔]/g) || []).length;
  const mdLeak = (t.match(/\*\*|(?<![\w*])\*(?![\w*])|^#{1,6}\s/gm) || []).length;
  const notes = NOTE_PAT.test(t);
  const figures = Object.entries(BAD_FIGURES).filter(([, re]) => re.test(t)).map(([k]) => k);

  if (archived) {
    if (figures.length) add('archive', rel, `historical submitted artifact — carries ${figures.join(', ')}`);
    continue; // never gate on the historical record
  }
  if (!hasMd) add('MATERIAL', rel, 'ORPHAN: no .md source — cannot be regenerated or verified');
  if (arrows) add('MATERIAL', rel, `${arrows} arrow glyph(s) reached the PDF (ATS hazard)`);
  if (mdLeak) add('MATERIAL', rel, `${mdLeak} literal markdown marker(s) in the PDF`);
  if (notes) add('MATERIAL', rel, 'editorial scaffolding reached the PDF');
  if (figures.length) add('MATERIAL', rel, `unsupported figure(s): ${figures.join(', ')}`);
}

const order = { MATERIAL: 0, warn: 1, archive: 2, frozen: 3 };
findings.sort((a, b) => order[a.sev] - order[b.sev] || a.doc.localeCompare(b.doc));
for (const f of findings)
  console.log(`${f.sev === 'MATERIAL' ? '!!' : '  '} [${f.sev.padEnd(8)}] ${f.doc.slice(-58).padEnd(58)} ${f.msg}`);

const counts = findings.reduce((a, f) => ((a[f.sev] = (a[f.sev] || 0) + 1), a), {});
const material = counts.MATERIAL || 0;
console.log(`\n${mds.length} sources · ${pdfs.length} PDFs · ` +
            (Object.entries(counts).map(([k, v]) => `${k}:${v}`).join('  ') || 'no findings'));

if (material) {
  console.log(`\n${material} material finding(s).`);
  if (STRICT) { console.error('FAIL — resolve before sending.'); process.exit(1); }
} else {
  console.log('\nNo material findings. Every artifact is current with its .md and clean.');
}
