# docx-track-changes-skill

Diffs two Markdown files and produces a `.docx` with **real Word
tracked-changes** (`w:ins`/`w:del` revision XML) — not colored or
struck-through text standing in for track changes. Word opens it in normal
review mode; Google Docs renders the same marks as native Suggesting-mode
edits on upload.

Built out of a recurring need: reviewing a contract/SOW maintained as
Markdown in git, where the actual reviewer (client, counsel, procurement)
wants to work in Word or Google Docs with accept/reject suggestions, not a
`git diff`.

See [`SKILL.md`](./SKILL.md) for exact scope (supported Markdown subset,
diffing strategy, known limitations) and a worked example — that file is
also the Claude Code skill descriptor, written so an agent can read it and
know when and how to use this.

## Requirements

- Python 3 with `python-docx` and `lxml`:
  `apt-get install python3-docx python3-lxml` (Debian/Ubuntu), or
  `pip install python-docx lxml`.
- `pandoc`, optional, only for `scripts/verify_redline.sh` (spot-checking
  output by round-tripping it back to text) — not needed to generate a
  redline.

## Usage

```
python3 scripts/build_redline.py OLD.md NEW.md OUT.docx --author "Your Name"
```

Typical case: diffing a file at two git refs.

```
git show main:path/to/doc.md > /tmp/old.md
cp path/to/doc.md /tmp/new.md   # working tree / branch version
python3 scripts/build_redline.py /tmp/old.md /tmp/new.md redline.docx --author "Your Name"
```

Upload `redline.docx` to Google Docs (File → Open, or drag into Drive) to
review it there with the same tracked changes.

## Verifying a change to the script

```
./scripts/verify_redline.sh
```

Builds a redline from `tests/fixtures/{old,new}.md` and asserts, via
pandoc, that accepting all changes reconstructs `new.md`'s content and
rejecting all changes reconstructs `old.md`'s content.

## Using this as a Claude Code skill

Clone this repo, then either:

- **Project-scoped:** symlink or copy it into the project's
  `.claude/skills/docx-track-changes/` so it's only offered in that repo.
- **User-scoped:** copy it into `~/.claude/skills/docx-track-changes/` so
  it's offered in every project.

Claude Code discovers `SKILL.md` in either location and lists the skill by
its frontmatter `name`/`description`; invoke it with `/docx-track-changes`
or let Claude reach for it when a redline is asked for. The script itself
has no Claude-specific dependency — it also runs standalone, as above.
