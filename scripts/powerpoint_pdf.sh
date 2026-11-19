#!/bin/bash
# usage: powerpoint_pdf.sh deck.pptx      -> deck.pdf beside it, written by PowerPoint itself (what the audience sees)
# macOS with Microsoft PowerPoint only.
# How: the deck is opened hidden in the background (`open -g -j`), that presentation (matched by file name, not
# "presentation 1", which is whatever is open in front) is saved as PDF inside the Office group container, where Office
# writes without a grant dialog, and the PDF is moved beside the deck.
# A deck of the same name that was already open is exported and left open; only a deck this script opened is closed
# (without saving). Other decks are not touched. A PowerPoint that was visible from the start stays visible: it has to
# be quit once before it launches hidden.
# No fallback to another renderer: a failure is reported as a failure. A run that works takes a few seconds; PowerPoint
# is given PPTX_TIMEOUT seconds (default 45) and then the step is killed and reported.
set -e
IN="$1"; [ -f "$IN" ] || { echo "usage: powerpoint_pdf.sh deck.pptx" >&2; exit 1; }
[ "$(uname)" = Darwin ] || { echo "powerpoint_pdf: needs PowerPoint on macOS" >&2; exit 2; }
[ -d "/Applications/Microsoft PowerPoint.app" ] || { echo "powerpoint_pdf: PowerPoint is not installed on this Mac" >&2; exit 2; }
D="$(cd "$(dirname "$IN")" && pwd)"; B="$(basename "${IN%.*}")"; IN="$D/$(basename "$IN")"; PDF="$D/$B.pdf"
rm -f "$PDF"
limit() { perl -e 'alarm shift; exec @ARGV or exit 127' "$@"; }   # limit <seconds> cmd...: killed at the limit (exit 142)
T="${PPTX_TIMEOUT:-45}"
BOX=~/Library/"Group Containers"/UBF8T346G9.Office/powerpoint-pdf-export
mkdir -p "$BOX"; rm -f "$BOX/$B.pdf"
WAS=no   # a deck of this name that is already open is exported and left open
pgrep -x "Microsoft PowerPoint" >/dev/null && limit 10 osascript -e 'tell application "Microsoft PowerPoint" to get name of presentations' 2>/dev/null | tr ',' '\n' | sed 's/^ //' | grep -qxF "$(basename "$IN")" && WAS=yes
open -g -j -a "Microsoft PowerPoint" "$IN"
RC=0; limit "$T" osascript - "$(basename "$IN")" "$BOX/$B.pdf" "$WAS" <<'EOF' || RC=$?
on run argv
  set want to item 1 of argv
  set hfs to (POSIX file (item 2 of argv)) as string
  tell application "Microsoft PowerPoint"
    set pres to missing value
    repeat 40 times
      repeat with i from 1 to (count of presentations)
        try
          if name of presentation i is want then set pres to presentation i
        end try
      end repeat
      if pres is not missing value then exit repeat
      delay 0.5
    end repeat
    if pres is missing value then error "PowerPoint did not open " & want
    save pres in hfs as save as PDF
    if (item 3 of argv) is "no" then close pres saving no
  end tell
end run
EOF
[ "$RC" = 142 ] && echo "powerpoint_pdf: PowerPoint gave no answer within $T s (locked screen or an open dialog?); PPTX_TIMEOUT=<seconds> allows more" >&2
if [ "$RC" != 0 ]; then   # do not leave our deck open
  [ "$WAS" = no ] && limit 10 osascript - "$(basename "$IN")" <<'EOF' >/dev/null 2>&1 || true
on run argv
  tell application "Microsoft PowerPoint" to close (first presentation whose name is (item 1 of argv)) saving no
end run
EOF
  exit 1
fi
[ -s "$BOX/$B.pdf" ] && mv "$BOX/$B.pdf" "$PDF"
[ -s "$PDF" ] || { echo "powerpoint_pdf: PowerPoint wrote no PDF for $IN" >&2; exit 1; }
echo "$PDF"
