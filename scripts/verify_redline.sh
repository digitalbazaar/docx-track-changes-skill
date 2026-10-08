#!/usr/bin/env bash
# Builds a redline from tests/fixtures/{old,new}.md and checks, via pandoc,
# that "accept all" reconstructs new.md's content and "reject all"
# reconstructs old.md's content. See SKILL.md's Check/eval section.
set -euo pipefail
cd "$(dirname "$0")/.."

OLD=tests/fixtures/old.md
NEW=tests/fixtures/new.md
OUT=$(mktemp --suffix=.docx)
trap 'rm -f "$OUT"' EXIT

python3 scripts/build_redline.py "$OLD" "$NEW" "$OUT" --author "CI" --date "2026-01-01T00:00:00Z"

if ! command -v pandoc >/dev/null; then
  echo "pandoc not found; cannot verify (pandoc is only needed for this check, not for generation)" >&2
  exit 1
fi

accepted=$(pandoc "$OUT" --track-changes=accept -t plain)
rejected=$(pandoc "$OUT" --track-changes=reject -t plain)

fail=0
assert_contains() {
  if ! grep -qF "$2" <<< "$1"; then
    echo "FAIL: expected to find '$2'" >&2
    fail=1
  fi
}
assert_absent() {
  if grep -qF "$2" <<< "$1"; then
    echo "FAIL: expected NOT to find '$2'" >&2
    fail=1
  fi
}

assert_contains "$accepted" "70,000"
assert_absent   "$accepted" "35,000"
assert_contains "$accepted" "Consulting"

assert_contains "$rejected" "35,000"
assert_absent   "$rejected" "70,000"
assert_absent   "$rejected" "Consulting"

if [ "$fail" -ne 0 ]; then
  echo "verify_redline: FAILED"
  exit 1
fi
echo "verify_redline: OK"
