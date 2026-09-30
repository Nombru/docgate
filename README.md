# docgate

Render Markdown to DOCX and PDF, and **refuse to ship an artifact that drifted from its source.**

Two small programs. Parsers that turn a Markdown file into a formatted document, and a gate that
audits every rendered artifact against the file it came from and exits non-zero if something is
wrong. The gate is the interesting half.

## Why it exists

A corpus of documents was maintained as Markdown and delivered as PDF. Each render was a
hand-transcription: someone opened a Word file and retyped the changes. Nobody checked the PDFs against
the Markdown, so nobody knew whether they matched.

An audit found two failures in the same corpus. One PDF was silently **missing three passages its
Markdown had**, because a page-fit trim was made at render time and never travelled back to the
source. Another PDF that had already been sent to a recipient carried **a line that appeared in no
Markdown file at all**. It was typed directly into the artifact months earlier, by someone who has
since forgotten doing it.

Neither was caught by review, because review reads the source and assumes the artifact matches.

The fix was structural. Parsers read the Markdown and generate the
artifacts, so there is exactly one copy of the words and a render cannot diverge from its source.
Then a gate makes the remaining failure modes cheap to detect.

## What the gate checks

| Finding | Meaning |
|---|---|
| `STALE` | The source is newer than the artifact beside it. The artifact is out of date. |
| `ORPHAN` | An artifact with no source. It cannot be regenerated or verified, and nothing records where it came from. |
| `ATS` | Arrow glyphs reached the output. They break applicant tracking systems (ATS) and other automated parsers downstream. |
| `MARKDOWN` | Literal `**`, `*` or `#` leaked into the rendered text. A parser bug that review will not see. |
| `NOTES` | Editorial scaffolding survived: `TODO`, `delete before sending`, a draft banner. |
| `FIGURES` | A claim a later accuracy pass retracted is still present in the artifact. |

That last one is the check a normal workflow is least likely to have. When a number is corrected in the source, the
already-rendered artifacts keep the old number, and nothing in a normal workflow notices. `FIGURES`
turns "we corrected that" into something a machine verifies.

```bash
node lib/check-drift.js --strict     # exit 1 on any material finding; use as a pre-ship hook
```

## Frozen documents

A document that has already been sent is a **historical record of what a recipient received**, and
regenerating it destroys that. List those in `frozen` and the gate reports their staleness without
failing on it. Expected drift and unexpected drift look different, and only one should block a
release.

## Configure

Everything corpus-specific lives in `docgate.config.json`, and nothing is hardcoded. Copy
`docgate.config.example.json` and edit:

```json
{
  "root": "./documents",
  "sourcePattern": "^(Report|Letter|Brief)",
  "frozen": ["2026-Q1 - Board Report"],
  "retractedFigures": { "pre-correction NPS": "NPS\\s*(of\\s*)?7[0-9]" },
  "author": { "name": "Your Name", "contact": ["you@example.com"] }
}
```

## Render

```bash
npm install                                        # the only dependency is `docx`
node lib/render-doc.js "path/to/Document.md"       # generic documents
node lib/render-resume.js "path/to/Resume.md"      # résumé layout
node lib/render-coverletter.js "path/to/Letter.md" # letter layout, anchored on the salutation
```

Each parser has a documented contract at the top of its file: which Markdown constructs map to
which document elements, and where it stops reading. The letter parser anchors on the greeting
line, so anything above it is treated as notes and never rendered. That is deliberate: it lets a
source file carry reviewer instructions that structurally cannot reach the output.

PDF conversion is left to LibreOffice:

```bash
soffice --headless --convert-to pdf --outdir "$(dirname "$F")" "$F"
```

## Not included

There is no test suite. The parsers were verified by rendering a real corpus and diffing extracted
PDF text against the sources, which is how the original drift was found, but that corpus is
private. If you adopt this, write the round-trip test (render, `pdftotext`, compare), because it
is the check that catches drift in the parsers themselves.
