#!/bin/sh
# SlideSize PowerPoint to PDF helper for macOS.
#
# Drives the copy of Microsoft PowerPoint installed on this Mac so that
# slidesize.com can batch convert presentations to PDF. Nothing is installed
# and nothing is sent anywhere. The page and this script exchange small JSON
# files inside the _slidesize folder of the output folder you chose:
#
#   queue/   job tickets written by the page
#   active/  the ticket being worked on
#   done/    one result per ticket, written by this script
#   in/      a temporary copy of each presentation, deleted after use
#   out/     PDFs the page still has to edit
#   ref/     PowerPoint's own pictures of slides, for the fidelity check
#
# To start it, open Terminal, type  sh  and a space, drag this file into the
# Terminal window and press Return. macOS will ask once whether Terminal may
# control Microsoft PowerPoint. Say yes. Close the Terminal window to stop.
#
# The script runs as two processes. This one, the watchdog, hands out
# tickets and keeps time. A second one, the worker, talks to PowerPoint
# through AppleScript. If PowerPoint stalls on a file the watchdog stops the
# worker, records the failure and carries on with the next file.
#
# Presentations you already have open are never closed or changed. Each file
# is converted from a temporary copy with its own name, which is closed
# without saving.

HELPER_VERSION="1.0.0"
PROTOCOL=1
APP_NAME="Microsoft PowerPoint"
ENGINE_NAME="${SLIDESIZE_HELPER_TEST_ENGINE_NAME:-Microsoft PowerPoint}"
PREFIX="slidesize-f"
TAB=$(printf '\t')
OSA="${SLIDESIZE_HELPER_OSASCRIPT:-osascript}"

# ------------------------------------------------------------------ small tools

now_iso() { date -u +%Y-%m-%dT%H:%M:%SZ; }
now_s() { date +%s; }

json_str() {
  printf '%s' "$1" | tr '\n\r\t' '   ' | sed 's/\\/\\\\/g; s/"/\\"/g'
}

write_atomic() {
  # $1 file, $2 content
  printf '%s\n' "$2" > "$1.$$.tmp" && mv -f "$1.$$.tmp" "$1"
}

num_field() {
  # $1 file, $2 key. Prints a number, true, false or null.
  [ -f "$1" ] || return 0
  sed -n 's/.*"'"$2"'":\([-0-9a-z.]*\).*/\1/p' "$1" 2>/dev/null | head -n 1
}

str_field() {
  [ -f "$1" ] || return 0
  sed -n 's/.*"'"$2"'":"\([^"]*\)".*/\1/p' "$1" 2>/dev/null | head -n 1
}

log() {
  printf '%s %s %s\n' "$(now_iso)" "$ROLE" "$1" >> "$ROOT/log/helper.log" 2>/dev/null
}

file_bytes() { wc -c < "$1" | tr -d ' '; }

int_list_json() {
  # "1 2 3" becomes [1,2,3]
  out=""
  for v in $1; do out="$out${out:+,}$v"; done
  printf '[%s]' "$out"
}

in_list() {
  for v in $1; do [ "$v" = "$2" ] && return 0; done
  return 1
}

count_list() { n=0; for v in $1; do n=$((n + 1)); done; printf '%s' "$n"; }

reverse_list() { out=""; for v in $1; do out="$v $out"; done; printf '%s' "$out"; }

# True while a child process is still running. A child that has exited but
# has not been collected yet does not count.
pid_running() {
  st=$(ps -o state= -p "$1" 2>/dev/null | tr -d ' ')
  case "$st" in ""|Z*) return 1 ;; esac
  return 0
}

# Runs a command with a time limit in seconds. Returns 124 when it ran out of time.
run_limited() {
  limit=$1; shift
  "$@" &
  cmd_pid=$!
  waited=0
  while pid_running "$cmd_pid"; do
    if [ "$waited" -ge "$limit" ]; then
      kill "$cmd_pid" 2>/dev/null
      wait "$cmd_pid" 2>/dev/null
      return 124
    fi
    sleep 1
    waited=$((waited + 1))
  done
  wait "$cmd_pid"
}

# ------------------------------------------------------------------ PowerPoint
#
# Each thing asked of PowerPoint is its own small AppleScript, written to a
# temporary folder when the helper starts. A script that this version of
# PowerPoint does not understand then fails on its own and is reported,
# without stopping the others.

write_scripts() {
  cat > "$SCRIPTS/version.applescript" <<'EOS'
on run argv
	with timeout of 120 seconds
		tell application "Microsoft PowerPoint"
			launch
			set v to version
		end tell
	end timeout
	return v
end run
EOS

  cat > "$SCRIPTS/others.applescript" <<'EOS'
on run argv
	set pre to item 1 of argv
	set n to 0
	with timeout of 20 seconds
		tell application "Microsoft PowerPoint"
			repeat with p in presentations
				if (name of p) does not start with pre then set n to n + 1
			end repeat
		end tell
	end timeout
	return n as text
end run
EOS

  cat > "$SCRIPTS/closestale.applescript" <<'EOS'
on run argv
	set pre to item 1 of argv
	with timeout of 60 seconds
		tell application "Microsoft PowerPoint"
			set stale to {}
			repeat with p in presentations
				if (name of p) starts with pre then set end of stale to (name of p)
			end repeat
			repeat with nm in stale
				close presentation (nm as text) saving no
			end repeat
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/open.applescript" <<'EOS'
on run argv
	set deckFile to POSIX file (item 1 of argv)
	with timeout of 7200 seconds
		tell application "Microsoft PowerPoint"
			launch
			open deckFile
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/facts.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	with timeout of 600 seconds
		tell application "Microsoft PowerPoint"
			set pres to presentation deckName
			set n to count of slides of pres
			set w to slide width of page setup of pres
			set h to slide height of page setup of pres
		end tell
	end timeout
	return (n as text) & tab & (w as text) & tab & (h as text)
end run
EOS

  cat > "$SCRIPTS/hidden.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	set found to ""
	with timeout of 600 seconds
		tell application "Microsoft PowerPoint"
			set pres to presentation deckName
			set n to count of slides of pres
			repeat with i from 1 to n
				if (hidden of slide show transition of slide i of pres) is true then set found to found & (i as text) & " "
			end repeat
		end tell
	end timeout
	return "ok " & found
end run
EOS

  cat > "$SCRIPTS/sethidden.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	set flag to ((item 2 of argv) is "1")
	with timeout of 600 seconds
		tell application "Microsoft PowerPoint"
			set pres to presentation deckName
			repeat with k from 3 to (count of argv)
				set hidden of slide show transition of slide ((item k of argv) as integer) of pres to flag
			end repeat
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/delete.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	with timeout of 600 seconds
		tell application "Microsoft PowerPoint"
			set pres to presentation deckName
			repeat with k from 2 to (count of argv)
				delete slide ((item k of argv) as integer) of pres
			end repeat
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/placeholder.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	set slideNo to (item 2 of argv) as integer
	set l to (item 3 of argv) as integer
	set t to (item 4 of argv) as integer
	set w to (item 5 of argv) as integer
	set h to (item 6 of argv) as integer
	set label to item 7 of argv
	with timeout of 120 seconds
		tell application "Microsoft PowerPoint"
			set sl to slide slideNo of presentation deckName
			set sh to make new shape at end of sl with properties {auto shape type:autoshape rectangle, left position:l, top:t, width:w, height:h}
			try
				set fore color of fill format of sh to {42, 42, 42}
			end try
			set content of text range of text frame of sh to label
			try
				set font size of font of text range of text frame of sh to 14
			end try
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/pdf.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	set outFile to POSIX file (item 2 of argv)
	with timeout of 7200 seconds
		tell application "Microsoft PowerPoint"
			save presentation deckName in outFile as save as PDF
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/png.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	set outFile to POSIX file (item 2 of argv)
	with timeout of 7200 seconds
		tell application "Microsoft PowerPoint"
			save presentation deckName in outFile as save as PNG
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/close.applescript" <<'EOS'
on run argv
	set deckName to item 1 of argv
	with timeout of 120 seconds
		tell application "Microsoft PowerPoint"
			close presentation deckName saving no
		end tell
	end timeout
	return "ok"
end run
EOS

  cat > "$SCRIPTS/quit.applescript" <<'EOS'
on run argv
	with timeout of 60 seconds
		tell application "Microsoft PowerPoint" to quit saving no
	end timeout
	return "ok"
end run
EOS
}

# Reads a ticket with the JavaScript that ships with macOS and prints one
# tab separated line per value, so the shell never has to parse JSON.
JXA_PARSE='function run(argv){var t=JSON.parse(argv[0]),o=t.options||{},L=[];function a(k,v){L.push(k+"\t"+String(v==null?"":v).replace(/[\t\r\n]+/g," "))}a("protocol",t.protocol);a("task",t.task||"convert");a("id",t.id);a("input",t.input);a("pdf",t.pdf);a("overwrite",t.overwrite?1:0);a("timeout",t.timeoutSec||600);a("source",t.sourceName);a("hidden",o.includeHidden?1:0);a("from",o.range?o.range[0]:"");a("to",o.range?o.range[1]:"");a("output",o.output||"slides");a("pdfa",o.pdfa?1:0);a("ref",o.reference?o.reference.mode:"");(o.placeholders||[]).forEach(function(p){L.push(["ph",p.slide,Math.round(p.left),Math.round(p.top),Math.round(p.width),Math.round(p.height),String(p.label||"").replace(/\t/g," ").replace(/\r?\n/g," / ")].join("\t"))});return L.join("\n")}'

JXA_FONTS='ObjC.import("AppKit");function run(){var m=$.NSFontManager.sharedFontManager,out={},lists=[m.availableFontFamilies,m.availableFonts];for(var k=0;k<lists.length;k++){var a=lists[k];for(var i=0;i<a.count;i++){out[ObjC.unwrap(a.objectAtIndex(i))]=1}}return JSON.stringify({platform:"macos",families:Object.keys(out).sort()})}'

JXA_PAGES='ObjC.import("Quartz");function run(argv){var d=$.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(argv[0]));return d.isNil()?"":String(d.pageCount)}'

osa() {
  verb=$1; shift
  "$OSA" "$SCRIPTS/$verb.applescript" "$@" 2> "$ERRF"
}

osa_error() {
  # the last line osascript printed, without its file and line prefix
  [ -f "$ERRF" ] || return 0
  tail -n 1 "$ERRF" | sed 's/^[^:]*:[0-9]*:[0-9]*: //'
}

pdf_pages() {
  "$OSA" -l JavaScript -e "$JXA_PAGES" "$1" 2>/dev/null
}

find_powerpoint() {
  PPT_APP=""
  for p in "/Applications/$APP_NAME.app" "$HOME/Applications/$APP_NAME.app"; do
    if [ -d "$p" ]; then PPT_APP="$p"; break; fi
  done
  if [ -z "$PPT_APP" ] && command -v mdfind >/dev/null 2>&1; then
    PPT_APP=$(mdfind "kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'" 2>/dev/null | head -n 1)
  fi
  PPT_VERSION=""; PPT_BUILD=""
  if [ -n "$PPT_APP" ]; then
    PPT_VERSION=$(defaults read "$PPT_APP/Contents/Info" CFBundleShortVersionString 2>/dev/null)
    PPT_BUILD=$(defaults read "$PPT_APP/Contents/Info" CFBundleVersion 2>/dev/null)
  fi
  if [ -n "$SLIDESIZE_HELPER_TEST_INSTALLED" ]; then PPT_APP="test"; PPT_VERSION="${SLIDESIZE_HELPER_TEST_VERSION:-0.0}"; PPT_BUILD="test"; fi
}

powerpoint_running() {
  if [ -n "$SLIDESIZE_HELPER_TEST_INSTALLED" ]; then return 1; fi
  pgrep -x "$APP_NAME" >/dev/null 2>&1
}

# ------------------------------------------------------------------ worker

write_worker() {
  # $1 state, $2 message
  write_atomic "$ROOT/worker.json" "{\"pid\":$$,\"state\":\"$1\",\"phase\":\"$PHASE\",\"id\":\"$CUR_ID\",\"heartbeat\":\"$(now_iso)\",\"others\":${OTHERS:-0},\"message\":\"$(json_str "$2")\"}"
}

set_phase() { PHASE=$1; write_worker busy ""; }

fail() {
  # $1 code, $2 message
  FAIL_CODE=$1; FAIL_MSG=$2
  return 1
}

parse_ticket() {
  PARSED=$("$OSA" -l JavaScript -e "$JXA_PARSE" "$(cat "$1")" 2> "$ERRF") || return 1
  T_PROTOCOL=""; T_TASK=""; T_ID=""; T_INPUT=""; T_PDF=""; T_OVERWRITE=0; T_TIMEOUT=600; T_SOURCE=""
  T_HIDDEN=0; T_FROM=""; T_TO=""; T_OUTPUT="slides"; T_PDFA=0; T_REF=""
  PH_FILE="$SCRIPTS/placeholders.$$"
  : > "$PH_FILE"
  while IFS="$TAB" read -r k v; do
    case "$k" in
      protocol) T_PROTOCOL=$v ;; task) T_TASK=$v ;; id) T_ID=$v ;; input) T_INPUT=$v ;; pdf) T_PDF=$v ;;
      overwrite) T_OVERWRITE=$v ;; timeout) T_TIMEOUT=$v ;; source) T_SOURCE=$v ;; hidden) T_HIDDEN=$v ;;
      from) T_FROM=$v ;; to) T_TO=$v ;; output) T_OUTPUT=$v ;; pdfa) T_PDFA=$v ;; ref) T_REF=$v ;;
      ph) printf '%s\n' "$v" >> "$PH_FILE" ;;
    esac
  done <<EOF
$PARSED
EOF
  [ "$T_PROTOCOL" = "$PROTOCOL" ]
}

# A path inside a ticket must stay inside the folders this helper owns.
resolve_paths() {
  IN_PATH=""; PDF_PATH=""
  case "$T_PDF" in
    ../*) b=${T_PDF#../}; case "$b" in */*|"") return 1 ;; esac; PDF_PATH="$OUTROOT/$b" ;;
    out/*) b=${T_PDF#out/}; case "$b" in */*|..*|"") return 1 ;; esac; PDF_PATH="$ROOT/out/$b" ;;
    *) return 1 ;;
  esac
  case "$PDF_PATH" in *.pdf|*.PDF) ;; *) return 1 ;; esac
  [ "$T_TASK" = "validate" ] && return 0
  case "$T_INPUT" in
    in/*) b=${T_INPUT#in/}; case "$b" in */*|..*|"") return 1 ;; esac; IN_PATH="$ROOT/in/$b"; DECK_NAME=$b ;;
    *) return 1 ;;
  esac
  return 0
}

export_pdf() {
  rm -f "$TMP_PDF"
  : > "$TMP_PDF"          # the file exists before PowerPoint is asked to write it, which avoids a permission prompt
  osa pdf "$DECK_NAME" "$TMP_PDF" >/dev/null || return 1
  [ -s "$TMP_PDF" ] || return 1
  PAGES=$(pdf_pages "$TMP_PDF")
  return 0
}

make_reference() {
  REF_SLIDES=""
  raw="$ROOT/ref/$T_ID-raw"
  rm -rf "$raw" "$raw".* "$ROOT/ref/$T_ID"
  set_phase reference
  run_limited $((120 + N * 3)) "$OSA" "$SCRIPTS/png.applescript" "$DECK_NAME" "$raw" >/dev/null 2> "$ERRF" || { REF_NOTE=$(osa_error); rm -rf "$raw" "$raw".*; return 1; }
  list="$SCRIPTS/png.$$"
  find "$ROOT/ref" -path "$raw*" -type f \( -name '*.png' -o -name '*.PNG' \) 2>/dev/null |
    awk '{ n=$0; sub(/.*\//, "", n); gsub(/[^0-9]/, " ", n); k=split(n, a, " "); print (k ? a[k] : 0) "\t" $0 }' | sort -n | cut -f2- > "$list"
  made=$(wc -l < "$list" | tr -d ' ')
  shown=""
  if [ "$made" -eq "$N" ]; then
    i=1; while [ "$i" -le "$N" ]; do shown="$shown $i"; i=$((i + 1)); done
  elif [ "$HID_KNOWN" = 1 ] && [ "$made" -eq $((N - $(count_list "$HID"))) ]; then
    i=1; while [ "$i" -le "$N" ]; do in_list "$HID" "$i" || shown="$shown $i"; i=$((i + 1)); done
  else
    REF_NOTE="PowerPoint wrote $made pictures for $N slides, so they could not be matched to slides."
    rm -rf "$raw" "$raw".*; rm -f "$list"
    return 1
  fi
  mkdir -p "$ROOT/ref/$T_ID"
  keep=$WANTED
  if [ "$T_REF" = "sample" ]; then
    c=$(count_list "$WANTED")
    if [ "$c" -gt 5 ]; then
      keep=""; k=0
      while [ "$k" -lt 5 ]; do
        idx=$(( (k * (c - 1) * 2 + 4) / 8 + 1 )); j=0
        for s in $WANTED; do j=$((j + 1)); if [ "$j" -eq "$idx" ]; then in_list "$keep" "$s" || keep="$keep $s"; fi; done
        k=$((k + 1))
      done
    fi
  fi
  for s in $shown; do
    IFS= read -r f || break
    if in_list "$keep" "$s"; then
      mv -f "$f" "$ROOT/ref/$T_ID/$(printf 'slide-%04d.png' "$s")" && REF_SLIDES="$REF_SLIDES $s"
    fi
  done < "$list"
  rm -rf "$raw" "$raw".*; rm -f "$list"
  return 0
}

convert() {
  START_S=$(now_s)
  N=0; W=0; H=0; HID=""; HID_KNOWN=0; WANTED=""; UNWANTED=""; REMOVED=""; PAGES=""; MAP_KNOWN=0
  NOT_APPLIED=""; PH_JSON=""; REF_SLIDES=""; REF_NOTE=""; OPEN_MS=0; EXPORT_MS=0; REF_MS=0; OPENED=0; HIDE_METHOD="none"
  [ -f "$IN_PATH" ] || { fail not-found "The temporary copy of the presentation is missing."; return 1; }
  if [ "$T_OUTPUT" = "notes" ] || [ "$T_PDFA" = 1 ]; then
    fail export-failed "Notes pages and PDF/A are not available from PowerPoint for Mac."; return 1
  fi
  TMP_PDF="$ROOT/out/$T_ID.export.pdf"

  set_phase opening
  osa open "$IN_PATH" >/dev/null || { fail open-failed "$(osa_error)"; return 1; }
  OPENED=1
  # PowerPoint answers the open before the presentation is ready to be asked about
  tries=0; FACTS=""
  while [ "$tries" -lt 40 ]; do
    FACTS=$(osa facts "$DECK_NAME") && [ -n "$FACTS" ] && break
    FACTS=""; tries=$((tries + 1)); sleep 0.5
  done
  [ -n "$FACTS" ] || { fail open-failed "PowerPoint opened the file but would not report its slides. $(osa_error)"; return 1; }
  N=$(printf '%s' "$FACTS" | cut -f1 | tr -cd '0-9')
  W=$(printf '%s' "$FACTS" | cut -f2 | tr ',' '.' | tr -cd '0-9.')
  H=$(printf '%s' "$FACTS" | cut -f3 | tr ',' '.' | tr -cd '0-9.')
  [ -n "$N" ] && [ "$N" -gt 0 ] || { fail no-slides "The presentation has no slides."; return 1; }
  OPEN_MS=$(( ($(now_s) - START_S) * 1000 ))

  h=$(osa hidden "$DECK_NAME")
  case "$h" in
    ok*) HID_KNOWN=1; HID=$(printf '%s' "${h#ok}" | tr -cd '0-9 ') ;;
    *) NOT_APPLIED="$NOT_APPLIED{\"option\":\"hidden\",\"reason\":\"PowerPoint for Mac would not say which slides are hidden. $(json_str "$(osa_error)")\"}," ;;
  esac

  from=1; to=$N
  if [ -n "$T_FROM" ]; then from=$T_FROM; to=$T_TO; fi
  [ "$from" -lt 1 ] && from=1
  [ "$to" -gt "$N" ] && to=$N
  OUT_OF_RANGE=""; UNHIDE=""
  i=1
  while [ "$i" -le "$N" ]; do
    if [ "$i" -lt "$from" ] || [ "$i" -gt "$to" ]; then
      UNWANTED="$UNWANTED $i"; in_list "$HID" "$i" || OUT_OF_RANGE="$OUT_OF_RANGE $i"
    elif in_list "$HID" "$i"; then
      if [ "$T_HIDDEN" = 1 ]; then WANTED="$WANTED $i"; UNHIDE="$UNHIDE $i"; else UNWANTED="$UNWANTED $i"; fi
    else
      WANTED="$WANTED $i"
    fi
    i=$((i + 1))
  done
  EXPECT=$(count_list "$WANTED")
  [ "$EXPECT" -gt 0 ] || { fail no-slides "No slides are left to export with this slide range and the hidden slides left out."; return 1; }

  while IFS="$TAB" read -r ps pl pt pw phh plabel; do
    [ -n "$ps" ] || continue
    if osa placeholder "$DECK_NAME" "$ps" "$pl" "$pt" "$pw" "$phh" "$plabel" >/dev/null; then
      PH_JSON="$PH_JSON{\"slide\":$ps,\"ok\":true},"
    else
      PH_JSON="$PH_JSON{\"slide\":$ps,\"ok\":false,\"error\":\"$(json_str "$(osa_error)")\"},"
    fi
  done < "$PH_FILE"

  if [ -n "$T_REF" ]; then
    r0=$(now_s)
    make_reference
    REF_MS=$(( ($(now_s) - r0) * 1000 ))
  fi

  set_phase exporting
  e0=$(now_s)
  # First ask PowerPoint to skip slides by hiding them, which changes nothing else.
  if [ -n "$UNHIDE" ]; then osa sethidden "$DECK_NAME" 0 $UNHIDE >/dev/null && HIDE_METHOD="hide"; fi
  if [ -n "$OUT_OF_RANGE" ]; then osa sethidden "$DECK_NAME" 1 $OUT_OF_RANGE >/dev/null && HIDE_METHOD="hide"; fi
  export_pdf || { fail export-failed "$(osa_error)"; return 1; }
  if [ -n "$PAGES" ] && [ "$PAGES" != "$EXPECT" ] && [ -n "$UNWANTED" ]; then
    # This PowerPoint exports hidden slides too. Take the unwanted slides out of the temporary copy and export again.
    if osa delete "$DECK_NAME" $(reverse_list "$UNWANTED") >/dev/null; then
      REMOVED=$UNWANTED; HIDE_METHOD="delete"
      export_pdf || { fail export-failed "$(osa_error)"; return 1; }
    fi
  fi
  if [ -n "$PAGES" ] && [ "$PAGES" = "$EXPECT" ]; then MAP_KNOWN=1; fi
  EXPORT_MS=$(( ($(now_s) - e0) * 1000 ))

  if [ -e "$PDF_PATH" ]; then
    if [ "$T_OVERWRITE" != 1 ]; then fail exists "The output file already exists: $(basename "$PDF_PATH")"; return 1; fi
    rm -f "$PDF_PATH"
  fi
  mv -f "$TMP_PDF" "$PDF_PATH" || { fail export-failed "The PDF could not be moved into the output folder."; return 1; }
  return 0
}

result_json() {
  # $1 ok (true/false), $2 status
  eng="{\"name\":\"$(json_str "$ENGINE_NAME")\",\"version\":\"$(json_str "$PPT_VERSION")\",\"build\":\"$(json_str "$PPT_BUILD")\",\"platform\":\"macos\"}"
  base="\"protocol\":$PROTOCOL,\"id\":\"$T_ID\",\"ok\":$1,\"status\":\"$2\",\"startedAt\":\"$STARTED\",\"finishedAt\":\"$(now_iso)\",\"engine\":$eng"
  if [ "$1" != true ]; then
    printf '{%s,"timeoutSec":%s,"error":{"code":"%s","message":"%s"}}' "$base" "${T_TIMEOUT:-600}" "$FAIL_CODE" "$(json_str "$FAIL_MSG")"
    return
  fi
  if [ "$T_TASK" = "validate" ]; then printf '{%s,"pdfa":null}' "$base"; return; fi
  hid=null; [ "$HID_KNOWN" = 1 ] && hid=$(int_list_json "$HID")
  map=null; [ "$MAP_KNOWN" = 1 ] && map=$(int_list_json "$WANTED")
  ref=null
  if [ -n "$REF_SLIDES" ]; then ref="{\"dir\":\"ref/$T_ID\",\"slides\":$(int_list_json "$REF_SLIDES"),\"pattern\":\"slide-%04d.png\"}"; fi
  [ -n "$T_REF" ] && [ -z "$REF_SLIDES" ] && ref="{\"dir\":null,\"slides\":[],\"note\":\"$(json_str "$REF_NOTE")\"}"
  total=$(( ($(now_s) - START_S) * 1000 ))
  printf '{%s,"facts":{"slides":%s,"hidden":%s,"slideWidthPt":%s,"slideHeightPt":%s,"fonts":null,"readOnly":false},' "$base" "$N" "$hid" "${W:-0}" "${H:-0}"
  printf '"export":{"method":"save as PDF","pages":%s,"pageMap":%s,"removedSlides":%s,"hiddenMethod":"%s","notApplied":[%s],"placeholders":[%s]},' \
    "${PAGES:-null}" "$map" "$(int_list_json "$REMOVED")" "$HIDE_METHOD" "${NOT_APPLIED%,}" "${PH_JSON%,}"
  printf '"pdf":{"bytes":%s},"reference":%s,"timing":{"openMs":%s,"exportMs":%s,"referenceMs":%s,"totalMs":%s}}' \
    "$(file_bytes "$PDF_PATH")" "$ref" "$OPEN_MS" "$EXPORT_MS" "$REF_MS" "$total"
}

run_ticket() {
  ticket=$1
  T_ID=$(basename "$ticket" .json)
  CUR_ID=$T_ID; PHASE=""; FAIL_CODE="convert-failed"; FAIL_MSG=""; OPENED=0; DECK_NAME=""; IN_PATH=""; T_TASK="convert"; T_TIMEOUT=600
  STARTED=$(now_iso)
  ok=false; status=failed
  if ! parse_ticket "$ticket"; then
    FAIL_CODE="bad-ticket"; FAIL_MSG="The ticket could not be read or comes from a different version of the page."
  elif T_ID=$(basename "$ticket" .json) && ! resolve_paths; then
    FAIL_CODE="bad-ticket"; FAIL_MSG="The ticket points outside the output folder."
  elif [ "$T_TASK" = "validate" ]; then
    ok=true; status=done
  else
    log "convert $T_ID $T_SOURCE"
    if convert; then ok=true; status=done; else log "failed $T_ID $FAIL_CODE $FAIL_MSG"; fi
  fi
  set_phase closing
  if [ "$OPENED" = 1 ]; then osa close "$DECK_NAME" >/dev/null 2>&1; fi
  T_ID=$(basename "$ticket" .json)
  [ -n "$IN_PATH" ] && rm -f "$IN_PATH"
  rm -f "$ROOT/out/$T_ID.export.pdf" "$PH_FILE"
  write_atomic "$ROOT/done/$T_ID.json" "$(result_json "$ok" "$status")"
  rm -f "$ticket"
  CUR_ID=""; PHASE=""
  OTHERS=$(osa others "$PREFIX" 2>/dev/null | tr -cd '0-9'); OTHERS=${OTHERS:-0}
  write_worker ready ""
}

worker_main() {
  ROLE="worker  "
  ERRF="$SCRIPTS/err.$$"
  PHASE=""; CUR_ID=""; OTHERS=0
  find_powerpoint
  write_worker starting ""
  if ! osa version >/dev/null; then
    msg=$(osa_error)
    log "could not start PowerPoint: $msg"
    write_worker error "PowerPoint could not be started. $msg"
    exit 2
  fi
  osa closestale "$PREFIX" >/dev/null 2>&1
  OTHERS=$(osa others "$PREFIX" 2>/dev/null | tr -cd '0-9'); OTHERS=${OTHERS:-0}
  if [ ! -f "$ROOT/fonts.json" ]; then
    f=$("$OSA" -l JavaScript -e "$JXA_FONTS" 2>/dev/null) && [ -n "$f" ] && write_atomic "$ROOT/fonts.json" "$f"
  fi
  write_worker ready ""
  beat=$(now_s)
  while :; do
    kill -0 "$PARENT_PID" 2>/dev/null || break
    [ -f "$ROOT/active/_stop" ] && break
    t=$(ls "$ROOT/active"/*.json 2>/dev/null | sort | head -n 1)
    if [ -n "$t" ]; then
      run_ticket "$t"
      beat=$(now_s)
    else
      if [ $(( $(now_s) - beat )) -ge 2 ]; then write_worker ready ""; beat=$(now_s); fi
      sleep 0.3
    fi
  done
  if [ "$OWNS" = 1 ]; then
    OTHERS=$(osa others "$PREFIX" 2>/dev/null | tr -cd '0-9')
    [ "${OTHERS:-1}" = 0 ] && osa quit >/dev/null 2>&1
  fi
  write_worker stopped ""
}

# ------------------------------------------------------------------ watchdog

find_root() {
  here=$(cd "$(dirname "$0")" 2>/dev/null && pwd)
  if [ -n "$SLIDESIZE_HELPER_ROOT" ]; then ROOT=$SLIDESIZE_HELPER_ROOT; return 0; fi
  if [ -f "$here/marker.json" ]; then ROOT=$here; return 0; fi
  if [ -f "$here/_slidesize/marker.json" ]; then ROOT="$here/_slidesize"; return 0; fi
  echo ""
  echo "This helper was started from outside the output folder."
  echo "Choose the output folder you picked on the SlideSize page."
  picked=$("$OSA" -e 'POSIX path of (choose folder with prompt "Choose the output folder you picked on the SlideSize page")' 2>/dev/null)
  picked=${picked%/}
  if [ -n "$picked" ] && [ -f "$picked/_slidesize/marker.json" ]; then ROOT="$picked/_slidesize"; return 0; fi
  [ -n "$picked" ] && echo "That folder has not been set up by the SlideSize page. Choose it on the page first."
  return 1
}

start_worker() {
  sh "$0" --worker "$ROOT" "$OWNS" "$SCRIPTS" "$$" &
  WORKER_PID=$!
  WORKER_STARTED=$(now_s)
}

stop_worker() {
  [ -n "$WORKER_PID" ] || return 0
  pkill -P "$WORKER_PID" 2>/dev/null
  kill "$WORKER_PID" 2>/dev/null
  wait "$WORKER_PID" 2>/dev/null
  WORKER_PID=""
}

worker_alive() { [ -n "$WORKER_PID" ] && pid_running "$WORKER_PID"; }

write_failure() {
  # $1 id, $2 status, $3 message
  write_atomic "$ROOT/done/$1.json" "{\"protocol\":$PROTOCOL,\"id\":\"$1\",\"ok\":false,\"status\":\"$2\",\"timeoutSec\":$ACTIVE_TIMEOUT,\"error\":{\"code\":\"$2\",\"message\":\"$(json_str "$3")\"},\"finishedAt\":\"$(now_iso)\",\"engine\":{\"name\":\"$(json_str "$ENGINE_NAME")\",\"version\":\"$(json_str "$PPT_VERSION")\",\"build\":\"$(json_str "$PPT_BUILD")\",\"platform\":\"macos\"}}"
}

# After a stalled worker has been stopped. Prints nothing when work can go on,
# or a sentence saying why it cannot.
recover() {
  powerpoint_running || return 0
  others=$(num_field "$ROOT/worker.json" others)
  if [ "$OWNS" = 1 ] && [ "${others:-0}" = 0 ]; then
    pkill -x "$APP_NAME" 2>/dev/null
    sleep 2
    return 0
  fi
  if run_limited 15 "$OSA" "$SCRIPTS/others.applescript" "$PREFIX" >/dev/null 2>&1; then return 0; fi
  echo "PowerPoint is not responding and it has your presentations open, so the helper will not close it. Answer any question PowerPoint is showing, or save your work and quit PowerPoint, then press Retry."
}

write_heartbeat() {
  cur=null
  if [ -n "$ACTIVE_ID" ]; then
    ph=""
    [ "$(str_field "$ROOT/worker.json" id)" = "$ACTIVE_ID" ] && ph=$(str_field "$ROOT/worker.json" phase)
    cur="{\"id\":\"$ACTIVE_ID\",\"phase\":\"$ph\",\"since\":\"$ACTIVE_SINCE_ISO\"}"
  fi
  inst=false; [ -n "$PPT_APP" ] && inst=true
  rb=false; [ "$RUNNING_BEFORE" = 1 ] && rb=true
  others=$(num_field "$ROOT/worker.json" others)
  shown=$STATE; [ -n "$PPT_APP" ] || shown=idle
  write_atomic "$ROOT/helper.json" "{\"protocol\":$PROTOCOL,\"helper\":\"$HELPER_VERSION\",\"platform\":\"macos\",\"os\":\"$(json_str "$OS_NAME")\",\"pid\":$$,\"heartbeat\":\"$(now_iso)\",\"state\":\"$shown\",\"message\":\"$(json_str "$MESSAGE")\",\"powerpoint\":{\"installed\":$inst,\"version\":\"$(json_str "$PPT_VERSION")\",\"build\":\"$(json_str "$PPT_BUILD")\",\"runningBefore\":$rb,\"userPresentations\":${others:-0}},\"caps\":{\"pdfa\":false,\"notes\":false,\"quality\":false,\"bitmapText\":false,\"tags\":false,\"docProps\":false,\"markup\":false,\"hidden\":true,\"range\":true,\"reference\":true,\"placeholders\":true},\"verapdf\":false,\"engine\":\"$(json_str "$ENGINE_NAME")\",\"current\":$cur,\"converted\":$CONVERTED,\"failed\":$FAILED}"
}

cleanup() {
  [ -n "$CLEANED" ] && return
  CLEANED=1
  : > "$ROOT/active/_stop" 2>/dev/null
  if worker_alive; then
    n=0; while worker_alive && [ "$n" -lt 8 ]; do sleep 1; n=$((n + 1)); done
    stop_worker
  fi
  rm -f "$ROOT/active/_stop"
  if [ -f "$ROOT/helper.json" ]; then
    STATE=stopped; write_heartbeat
  fi
  rmdir "$ROOT/helper.lock" 2>/dev/null
  [ -n "$SCRIPTS" ] && rm -rf "$SCRIPTS"
  log "stop"
  echo ""
  echo "Helper stopped."
}

watchdog_main() {
  ROLE="watchdog"
  if ! find_root; then
    echo ""
    echo "No SlideSize output folder was found. Go back to the page, choose an output folder, then start this helper again."
    exit 1
  fi
  ROOT=$(cd "$ROOT" && pwd)
  OUTROOT=$(dirname "$ROOT")
  for d in queue active done in out ref log; do mkdir -p "$ROOT/$d"; done

  # one helper per folder. A lock left by a helper that is no longer running is cleared.
  if ! mkdir "$ROOT/helper.lock" 2>/dev/null; then
    old=$(num_field "$ROOT/helper.json" pid)
    if [ -n "$old" ] && kill -0 "$old" 2>/dev/null; then
      echo ""
      echo "A helper is already running for this folder. Use that one, or close its window first."
      exit 1
    fi
  fi

  for f in "$ROOT/active"/*; do
    [ -e "$f" ] || continue
    case "$f" in */_stop) rm -f "$f" ;; *) mv -f "$f" "$ROOT/queue/" 2>/dev/null ;; esac
  done
  rm -f "$ROOT/worker.json" "$ROOT/fonts.json"

  SCRIPTS=$(mktemp -d "${TMPDIR:-/tmp}/slidesize-helper.XXXXXX") || exit 1
  ERRF="$SCRIPTS/err.$$"
  write_scripts
  find_powerpoint
  OS_NAME="macOS $(sw_vers -productVersion 2>/dev/null)"
  RUNNING_BEFORE=0; powerpoint_running && RUNNING_BEFORE=1
  OWNS=1; [ "$RUNNING_BEFORE" = 1 ] && OWNS=0
  LAST_SEQ=$(num_field "$ROOT/control.json" seq); LAST_SEQ=${LAST_SEQ:-0}

  echo ""
  echo "SlideSize PowerPoint to PDF helper $HELPER_VERSION"
  echo "Folder   $OUTROOT"
  if [ -z "$PPT_APP" ]; then
    echo ""
    echo "Microsoft PowerPoint was not found on this Mac."
    echo "Microsoft PowerPoint must be installed for PowerPoint-based conversion."
  else
    [ "$RUNNING_BEFORE" = 1 ] && echo "PowerPoint is already open. Your presentations will be left alone."
    echo "PowerPoint will show each presentation briefly while it converts it."
    echo "Waiting for the SlideSize page. Leave this window open. Close it to stop."
  fi
  log "start $HELPER_VERSION powerpoint=$PPT_VERSION runningBefore=$RUNNING_BEFORE"

  trap cleanup EXIT
  trap 'exit 0' INT TERM HUP

  WORKER_PID=""; STATE=starting; MESSAGE=""; ACTIVE_ID=""; ACTIVE_SINCE=0; ACTIVE_SINCE_ISO=""; ACTIVE_TIMEOUT=600
  CONVERTED=0; FAILED=0; LAST_BEAT=0; STOP=0

  while [ "$STOP" = 0 ]; do
    NOW=$(now_s)

    # the page can ask to stop, or to give up on the file in hand
    ABORT=""
    seq=$(num_field "$ROOT/control.json" seq)
    if [ -n "$seq" ] && [ "$seq" -gt "$LAST_SEQ" ] 2>/dev/null; then
      LAST_SEQ=$seq
      [ "$(num_field "$ROOT/control.json" stop)" = true ] && STOP=1
      ABORT=$(str_field "$ROOT/control.json" abort)
      if [ "$(num_field "$ROOT/control.json" retry)" = true ] && [ "$STATE" = blocked ]; then STATE=starting; MESSAGE=""; fi
    fi
    [ "$STOP" = 1 ] && break

    if [ -n "$PPT_APP" ] && [ "$STATE" != blocked ]; then
      if ! worker_alive; then
        if [ -n "$WORKER_PID" ] && [ -n "$ACTIVE_ID" ]; then
          write_failure "$ACTIVE_ID" failed "The helper process that talks to PowerPoint stopped unexpectedly while converting this file."
          rm -f "$ROOT/active/$ACTIVE_ID.json"
          FAILED=$((FAILED + 1)); ACTIVE_ID=""
        fi
        if [ -n "$WORKER_PID" ] && [ "$(str_field "$ROOT/worker.json" state)" = error ]; then
          STATE=blocked; MESSAGE=$(str_field "$ROOT/worker.json" message); WORKER_PID=""
        else
          start_worker; STATE=starting
        fi
      fi
    fi

    if worker_alive; then
      ws=$(str_field "$ROOT/worker.json" state)
      if [ "$ws" = ready ] || [ "$ws" = busy ]; then [ "$STATE" = starting ] && STATE=idle; fi
      if [ "$STATE" = starting ] && [ $((NOW - WORKER_STARTED)) -gt 150 ]; then
        stop_worker
        MESSAGE=$(recover)
        [ -n "$MESSAGE" ] || MESSAGE="PowerPoint did not start within 150 seconds. If macOS asked whether Terminal may control Microsoft PowerPoint, allow it, then press Retry."
        STATE=blocked
      fi
    fi

    if [ -n "$ACTIVE_ID" ] && [ -f "$ROOT/done/$ACTIVE_ID.json" ]; then
      if [ "$(num_field "$ROOT/done/$ACTIVE_ID.json" ok)" = true ]; then
        CONVERTED=$((CONVERTED + 1)); echo "  done    $ACTIVE_ID  $((NOW - ACTIVE_SINCE)) s"
      else
        FAILED=$((FAILED + 1)); echo "  FAILED  $ACTIVE_ID  $(str_field "$ROOT/done/$ACTIVE_ID.json" message)"
      fi
      ACTIVE_ID=""
      [ "$STATE" != blocked ] && STATE=idle
    fi

    if [ -n "$ACTIVE_ID" ]; then
      late=0; [ $((NOW - ACTIVE_SINCE)) -gt "$ACTIVE_TIMEOUT" ] && late=1
      if [ "$late" = 1 ] || [ "$ABORT" = "$ACTIVE_ID" ]; then
        stop_worker
        if [ "$late" = 1 ]; then st=timeout; text="PowerPoint did not finish within $ACTIVE_TIMEOUT seconds. It may be showing a dialog, such as a request to grant access to a folder."
        else st=cancelled; text="Skipped from the page while it was converting."; fi
        write_failure "$ACTIVE_ID" "$st" "$text"
        rm -f "$ROOT/active/$ACTIVE_ID.json" "$ROOT/out/$ACTIVE_ID.export.pdf"
        echo "  $st $ACTIVE_ID"; log "$st $ACTIVE_ID"
        FAILED=$((FAILED + 1)); ACTIVE_ID=""
        MESSAGE=$(recover)
        if [ -n "$MESSAGE" ]; then STATE=blocked; else STATE=starting; fi
      fi
    elif [ -n "$ABORT" ]; then
      case "$ABORT" in *[!A-Za-z0-9_-]*) ;; *) rm -f "$ROOT/queue/$ABORT.json" ;; esac
    fi

    if [ -z "$ACTIVE_ID" ] && [ "$STATE" = idle ] && worker_alive; then
      next=$(ls "$ROOT/queue"/*.json 2>/dev/null | sort | head -n 1)
      if [ -n "$next" ]; then
        id=$(basename "$next" .json)
        case "$id" in
          *[!A-Za-z0-9_-]*) rm -f "$next" ;;
          *)
            tmo=$(num_field "$next" timeoutSec); ACTIVE_TIMEOUT=${tmo:-600}
            case "$ACTIVE_TIMEOUT" in *[!0-9]*|"") ACTIVE_TIMEOUT=600 ;; esac
            if mv -f "$next" "$ROOT/active/$id.json" 2>/dev/null; then
              ACTIVE_ID=$id; ACTIVE_SINCE=$NOW; ACTIVE_SINCE_ISO=$(now_iso); STATE=working
              echo "  convert  $id  $(str_field "$ROOT/active/$id.json" sourceName)"
            fi
            ;;
        esac
      fi
    fi
    [ -n "$ACTIVE_ID" ] && [ "$STATE" = idle ] && STATE=working

    if [ $((NOW - LAST_BEAT)) -ge 2 ]; then LAST_BEAT=$NOW; write_heartbeat; fi
    sleep 0.3
  done
  exit 0
}

if [ "$1" = "--worker" ]; then
  ROOT=$2; OWNS=$3; SCRIPTS=$4; PARENT_PID=$5
  OUTROOT=$(dirname "$ROOT")
  worker_main
else
  watchdog_main
fi
