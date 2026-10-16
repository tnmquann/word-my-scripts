#!/bin/bash
# Install (sideload) the MY_SCRIPTS Word add-in on macOS for the current user.
# UNVERIFIED: written without a Mac to test on. Report problems on GitHub Issues.
#
#   curl -fsSL {{BASE_URL}}install.sh | bash
#   bash install.sh [--clear-cache] [--no-desktop-tools]
#
# - Manifest -> ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/<id>.manifest.xml
#   (Word's sideload folder; the add-in then appears under Home > Add-ins > DEVELOPER ADD-INS).
# - MY_SCRIPTS_Desktop.dotm (MathType tools, VBA) -> Word's Startup folder in the Office group container.
# - No auto-updater on macOS: code updates are live; re-run this script after a version with new buttons.
set -euo pipefail

BASE_URL='{{BASE_URL}}'
CLEAR_CACHE=0
DESKTOP_TOOLS=1
for a in "$@"; do
  case "$a" in
    --clear-cache) CLEAR_CACHE=1 ;;
    --no-desktop-tools) DESKTOP_TOOLS=0 ;;
    *) echo "Unknown option: $a" >&2; exit 2 ;;
  esac
done
case "$BASE_URL" in *'{{'*) echo "install.sh was not built (BASE_URL missing)." >&2; exit 1 ;; esac
[ "$(uname -s)" = "Darwin" ] || { echo "This installer is for macOS. On Windows use install.ps1." >&2; exit 1; }

step() { printf '\033[36m==> %s\033[0m\n' "$1"; }
HERE="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" 2>/dev/null && pwd || echo "")"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

# 1. Manifest: next to this script (release zip) or from the site
if [ -n "$HERE" ] && [ -f "$HERE/manifest.xml" ]; then
  cp "$HERE/manifest.xml" "$TMP/manifest.xml"
else
  step "Downloading manifest: ${BASE_URL}manifest.xml"
  curl -fsSL "${BASE_URL}manifest.xml" -o "$TMP/manifest.xml"
fi
ID="$(sed -n 's:.*<Id>\(.*\)</Id>.*:\1:p' "$TMP/manifest.xml" | head -n 1)"
[ -n "$ID" ] || { echo "Invalid manifest (no <Id>)." >&2; exit 1; }

WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
mkdir -p "$WEF"
cp "$TMP/manifest.xml" "$WEF/$ID.manifest.xml"
step "Manifest installed: $WEF/$ID.manifest.xml"

# 2. Desktop tools (MathType): Startup folder of Word (names are localized on some systems)
if [ "$DESKTOP_TOOLS" = 1 ]; then
  GC="$HOME/Library/Group Containers/UBF8T346G9.Office"
  STARTUP="$GC/User Content.localized/Startup.localized/Word"
  [ -d "$GC/User Content.localized" ] || STARTUP="$GC/User Content/Startup/Word"
  if mkdir -p "$STARTUP" 2>/dev/null; then
    if [ -n "$HERE" ] && [ -f "$HERE/MY_SCRIPTS_Desktop.dotm" ]; then
      cp "$HERE/MY_SCRIPTS_Desktop.dotm" "$STARTUP/"
    else
      curl -fsSL "${BASE_URL}MY_SCRIPTS_Desktop.dotm" -o "$STARTUP/MY_SCRIPTS_Desktop.dotm"
    fi
    step "Desktop tools (MathType) installed: $STARTUP/MY_SCRIPTS_Desktop.dotm"
  else
    echo "Warning: could not create $STARTUP; desktop tools skipped." >&2
  fi
fi

# 3. Optional: clear the Office add-in cache (Word must be closed)
if [ "$CLEAR_CACHE" = 1 ]; then
  if pgrep -x "Microsoft Word" >/dev/null; then
    echo "Warning: Word is running; quit Word and re-run with --clear-cache." >&2
  else
    rm -rf "$HOME/Library/Containers/com.microsoft.Word/Data/Library/Caches/"* \
           "$HOME/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef/"* 2>/dev/null || true
    step "Cleared Office add-in cache"
  fi
fi

printf '\n\033[32mDone.\033[0m Quit Word completely (Cmd+Q) and open it again.\n'
echo "If the MY_SCRIPTS tab is missing: Home > Add-ins (or Insert > Add-ins) > Developer Add-ins > MY_SCRIPTS (may be needed after each restart)."
echo "Desktop tools: tab MY_SCRIPTS+ (MathType tools need MathType for Mac; allow macros if Word asks)."
