#!/usr/bin/env bash
#
# Uninstalls W365 CLI from macOS.
#
# Removes ~/.local/bin/w365cli. The PATH line added to your shell profile by install.sh is left
# in place (harmless once the binary is gone) unless you pass --purge-path.
#
# Usage:
#   curl -fsSL https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/uninstall.sh | bash
#
# Options (pass after the script when running locally, not supported via curl | bash):
#   --purge-path   Also remove the PATH line added to your shell profile.

set -euo pipefail

INSTALL_DIR="$HOME/.local/bin"
BIN_NAME="w365cli"
DEST_BIN="$INSTALL_DIR/$BIN_NAME"
PURGE_PATH=0

for arg in "$@"; do
  case "$arg" in
    --purge-path) PURGE_PATH=1 ;;
    *) echo "Unknown option: $arg" >&2; exit 1 ;;
  esac
done

echo ""
echo "Uninstalling W365 CLI..."

if [ -f "$DEST_BIN" ]; then
  rm -f "$DEST_BIN"
  echo "Removed $DEST_BIN"
else
  echo "W365 CLI doesn't appear to be installed at $DEST_BIN"
fi

if [ "$PURGE_PATH" -eq 1 ]; then
  path_line='export PATH="$HOME/.local/bin:$PATH"'
  for rc_file in "$HOME/.zshrc" "$HOME/.bash_profile" "$HOME/.profile"; do
    if [ -f "$rc_file" ] && grep -qF "$path_line" "$rc_file" 2>/dev/null; then
      # Remove the marker comment line and the export line added by install.sh, leaving
      # everything else in the file untouched.
      tmp_file="$(mktemp)"
      grep -vF -e "$path_line" -e "# Added by W365 CLI installer" "$rc_file" > "$tmp_file"
      mv "$tmp_file" "$rc_file"
      echo "Removed PATH entry from $rc_file"
    fi
  done
fi

echo ""
echo "W365 CLI uninstalled."
echo ""
