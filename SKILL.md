---
name: docx-track-changes
description: "use when asked to produce a Word/Google Docs redline (tracked changes, suggesting-mode) between two versions of a Markdown document"
version: 0.1.0
model_req: none
owners: [BigBlueHat]
core: false
---

# Docx Track Changes

## What it does
Diffs two Markdown files and produces a `.docx` carrying real Word
tracked-changes XML (`w:ins` / `w:del` revision marks with author and
timestamp) — not colored or struck-through text standing in for track
changes. Word opens it in normal track-changes review mode; Google Docs
(on upload) renders the same marks as native Suggesting-mode edits, so a
reviewer can accept/reject individual changes rather than read a static
diff.

Use it whenever someone needs to review a Markdown-authored document's
changes (a contract, SOW, policy doc, spec) in Word or Google Docs instead
of as raw Markdown/git diff. Point it at the "before" and "after" Markdown
files — e.g. a file at a base git ref and the same file on a branch — and
it produces the redline directly; it does not require either file to go
through an intermediate Word conversion first.

Run it as:
```
python3 scripts/build_redline.py OLD.md NEW.md OUT.docx --author "Your Org"
```
`--author` names the revision author shown in Word's and Google Docs'
reviewer list, and is **required** — pass it, or set the `REDLINE_AUTHOR`
environment variable once (e.g. in your shell profile or CI config) so you
don't have to pass it on every call. There is deliberately no built-in
default: this is a shared public tool, used by whoever clones it and by
whatever agent runs it, so it has no org name or agent name of its own to
reach for. A generic placeholder (an AI's name, a made-up org) ending up
in someone's reviewed document unannounced is worse than the tool simply
refusing to guess — pass `--author` or set `REDLINE_AUTHOR`, or the CLI
exits with an error telling you to. `--date` pins an ISO-8601 UTC timestamp
instead of "now", useful for reproducible output.

Requires `python3-docx` and `python3-lxml` (`apt-get install python3-docx
python3-lxml` on Debian/Ubuntu, or `pip install python-docx lxml`). `pandoc`
is optional — only needed to spot-check output (see below), not to generate it.

### Diffing strategy
- Parses each file into a flat sequence of blocks: headings, `- ` list items
  (one level of nesting), pipe-table rows, and blank-line-separated
  paragraphs.
- Diffs that block sequence with `difflib.SequenceMatcher`. Unchanged blocks
  render plainly.
- A single old block replaced by a single new block of the same type gets a
  **word-level** diff (e.g. a changed dollar figure shows as just that word
  struck and replaced, not the whole sentence) — this is the common case for
  small edits (numbers, renamed terms, reworded clauses).
- A table that changed gets a **row-level** diff instead (rows map to table
  rows, not blocks), so an added/removed SKU or line item shows as a tracked
  row insertion/deletion rather than retyping the whole table.
- Anything else in a replace range (a block's type changed, or many old
  blocks became many new blocks — e.g. a section was restructured, not just
  reworded) falls back to **deleting all the old blocks and inserting all
  the new ones** as whole units. This is intentional: cross-matching
  unrelated bullets word-for-word to shave the diff size produces a
  confusing redline for genuine restructuring. A full delete-then-insert of
  the changed section is the more honest and more reviewable result when
  the content isn't a simple reword.

## What it does NOT do
- Not a general CommonMark/GFM engine. Supported inline syntax: `**bold**`,
  `_italic_` / `*italic*`, `~strike~` / `~~strike~~`, and `` `code` `` spans
  (rendered literal, not monospaced). Each requires a matching closing
  delimiter later in the same block — an unpaired `_`, `*`, or `~` (a
  snake_case word, a stray tilde) is treated as a literal character rather
  than toggling formatting on for the rest of the block. Delimiters nest
  correctly when one opens inside another and closes before it (`**bold
  _and italic_**`), but two separate spans placed side by side are not
  re-merged. Links, images, footnotes, numbered lists, and nested lists
  beyond one level are not parsed — they pass through as literal text in a
  plain paragraph rather than causing an error, so output stays readable
  even outside the supported subset, just without that element's intended
  structure.
- Does not silently drop table cells. A row whose cell count doesn't match
  its own table's header raises an error naming the row, rather than
  truncating the extra cells — a ragged/hand-edited markdown table is a
  mistake worth surfacing, not swallowing. Old and new table versions are
  allowed to have genuinely different column counts (a column added or
  removed is a valid edit); only a row disagreeing with *its own* header is
  an error.
- Does not detect moved content as "moved." If a paragraph relocates to a
  different section with identical text, `SequenceMatcher`'s longest-match
  behavior may mark it unchanged even though its surrounding context
  changed — a known limitation of sequence-based (not semantic) diffing,
  shared by most "compare documents" tools.
- Does not open, upload, or email the resulting file anywhere. It writes a
  local `.docx`; getting it into Google Docs (or wherever it needs review)
  is a separate, manual step.
- Does not accept a `.docx` as input on either side — both inputs must be
  Markdown. Converting an existing Word doc to Markdown first is out of
  scope for this skill.

## Worked example
Input: `old.md` contains `**Fixed Fee: $35,000**`; `new.md` contains the
same sentence with `**Fixed Fee: $70,000**`; a table in `new.md` has one
added row not present in `old.md`.

Output: a `.docx` where opening Track Changes review shows `~~$35,000~~
$70,000` as a single word-level revision (not the whole sentence struck),
and the added table row is marked as a tracked row insertion. Running
`python3 scripts/build_redline.py old.md new.md out.docx --author "Jane
Doe"` produces this directly.

## Check / eval
No LLM is in this skill's execution path — it's a deterministic diff/render
script, so there's no model-output drift to eval. Correctness is checked by
round-tripping the output through `pandoc`:
```
pandoc out.docx --track-changes=accept -t plain   # must read back as new.md's content
pandoc out.docx --track-changes=reject -t plain   # must read back as old.md's content
```
Both should reconstruct their respective source's prose (modulo Markdown
syntax characters, which plain-text extraction drops). A regression that
corrupts tracked-change nesting or loses content will show up as a mismatch
in one or both reconstructions. `scripts/verify_redline.sh` runs both
checks and asserts a handful of expected substrings from each file.

This check has a known blind spot: pandoc's own `--track-changes=accept`
doesn't fully honor row-level table deletion (`w:trPr`/`w:del`) — it blanks
a deleted row's cell text but still renders an empty row, even though the
underlying XML is spec-correct and matches the already-verified row
*insertion* pattern exactly. Table-row deletion was confirmed correct by
direct XML inspection, not by this round-trip; a real Word or Google Docs
accept-all is the authoritative check if that path changes again.
