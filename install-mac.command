#!/bin/bash
# ---------------------------------------------------------------------------
# Client Export Tool - macOS installer
#
# Run from Terminal: type  bash  and a space, drag this file into the window,
# press Return. (Running it through bash avoids Gatekeeper blocking an
# unsigned script.) It:
#   1. copies ClientExportTool.xlam into Excel's Add-Ins folder, where it is
#      listed automatically under Tools > Excel Add-ins
#   2. removes the macOS quarantine flag added when the file was downloaded
# Run it again at any time to update to a newer version.
#
# Keep this file with LF line endings (see .gitattributes).
# ---------------------------------------------------------------------------
set -u

ADDIN_FILE="ClientExportTool.xlam"
OFFICE="$HOME/Library/Group Containers/UBF8T346G9.Office"

ok()   { printf '\033[32m%s\033[0m\n' "$1"; }
warn() { printf '\033[33m%s\033[0m\n' "$1"; }
fail() { printf '\n\033[31mERROR:\033[0m %s\n' "$1"; exit 1; }

printf '\033[36mClient Export Tool - macOS installer\033[0m\n\n'

here="$(cd "$(dirname "$0")" && pwd)"

# 1. Find the add-in next to this installer ---------------------------------
src=""
count=0
for f in "$here"/ClientExportTool*.xlam; do
    [ -f "$f" ] || continue
    count=$((count + 1))
    if [ -z "$src" ] || [ "$f" -nt "$src" ]; then src="$f"; fi
done
[ -n "$src" ] || fail "Couldn't find $ADDIN_FILE in $here. Keep this installer in the same folder as the add-in."
[ "$count" -gt 1 ] && warn "More than one add-in file is in this folder; using the newest: $(basename "$src")"

# 2. Find Excel's Add-Ins folder --------------------------------------------
[ -d "$OFFICE" ] || fail "Couldn't find Microsoft Office's settings folder. Open Excel once, quit it, then run this installer again."

dest_dir=""
for d in "User Content.localized/Add-Ins.localized" "User Content/Add-Ins" \
         "User Content.localized/Add-Ins" "User Content/Add-Ins.localized"; do
    if [ -d "$OFFICE/$d" ]; then dest_dir="$OFFICE/$d"; break; fi
done
if [ -z "$dest_dir" ]; then
    if [ -d "$OFFICE/User Content.localized" ]; then
        dest_dir="$OFFICE/User Content.localized/Add-Ins.localized"
    else
        dest_dir="$OFFICE/User Content/Add-Ins"
    fi
    mkdir -p "$dest_dir" 2>/dev/null || fail "Couldn't create $dest_dir. If macOS asked whether Terminal may access data from other apps, choose Allow and run this again."
fi
dest="$dest_dir/$ADDIN_FILE"

# 3. Excel must be closed -----------------------------------------------------
while pgrep -xq "Microsoft Excel"; do
    warn "Excel is running. Save your work, quit Excel (Cmd+Q), then press Return here. (Ctrl+C to cancel.)"
    read -r _
done

# 4. Copy and remove quarantine -----------------------------------------------
if [ "$src" != "$dest" ]; then
    cp -f "$src" "$dest" 2>/dev/null || fail "Couldn't copy the add-in to $dest_dir. If macOS asked whether Terminal may access data from other apps, choose Allow and run this again."
fi
xattr -d com.apple.quarantine "$dest" 2>/dev/null
if xattr -p com.apple.quarantine "$dest" >/dev/null 2>&1; then
    fail "macOS still has the add-in quarantined. Run:  xattr -d com.apple.quarantine \"$dest\""
fi
ok "Installed: $dest"

echo
ok "Done!"
echo "First install only: open Excel, choose Tools > Excel Add-ins..., tick ClientExportTool, click OK."
echo "(Updating? If it's already ticked, there's nothing else to do.)"
echo "The button is under Review > Client Export > Create Client Export."
echo
read -r -p "Open Excel now? [Y/n] " ans
case "$ans" in
    [nN]*) ;;
    *) open -a "Microsoft Excel" 2>/dev/null || warn "Couldn't open Excel automatically; open it from Applications." ;;
esac
exit 0
