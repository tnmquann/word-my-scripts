#!/bin/bash
# Remove the MY_SCRIPTS Word add-in and its desktop tools on macOS (current user).
# UNVERIFIED: written without a Mac to test on.
#   curl -fsSL {{BASE_URL}}uninstall.sh | bash
set -euo pipefail

ID='{{ID}}'
case "$ID" in *'{{'*) echo "uninstall.sh was not built (ID missing)." >&2; exit 1 ;; esac

WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
if [ -f "$WEF/$ID.manifest.xml" ]; then
  rm -f "$WEF/$ID.manifest.xml"
  echo "==> Removed manifest $ID"
else
  echo "==> Manifest $ID was not installed"
fi

GC="$HOME/Library/Group Containers/UBF8T346G9.Office"
for d in "$GC/User Content.localized/Startup.localized/Word" "$GC/User Content/Startup/Word"; do
  if [ -f "$d/MY_SCRIPTS_Desktop.dotm" ]; then
    rm -f "$d/MY_SCRIPTS_Desktop.dotm"
    echo "==> Removed desktop tools ($d)"
  fi
done
echo "Done. Quit and reopen Word to remove the MY_SCRIPTS tabs."
