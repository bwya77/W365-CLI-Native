#!/usr/bin/env bash
#
# Installs W365 CLI for macOS or Linux.
#
# Downloads the latest W365 CLI build from GitHub (signed & notarized on macOS), installs it to
# ~/.local/bin/w365cli (no sudo required), and adds ~/.local/bin to your PATH if it isn't
# already there.
#
# Usage:
#   curl -fsSL https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/install.sh | bash
#
# Options (pass after the script when running locally, not supported via curl | bash):
#   --no-path   Install without adding ~/.local/bin to your PATH.

set -euo pipefail

REPO="bwya77/W365-CLI-Native"
INSTALL_DIR="$HOME/.local/bin"
BIN_NAME="w365cli"
NO_PATH=0

for arg in "$@"; do
  case "$arg" in
    --no-path) NO_PATH=1 ;;
    *) echo "Unknown option: $arg" >&2; exit 1 ;;
  esac
done

# --- OS and architecture detection ------------------------------------------
uname_os="$(uname -s)"
case "$uname_os" in
  Darwin) os="macos" ;;
  Linux) os="linux" ;;
  *)
    echo "W365 CLI supports macOS and Linux. Detected: $uname_os" >&2
    echo "For Windows, use install.ps1 instead: https://github.com/$REPO" >&2
    exit 1
    ;;
esac

uname_arch="$(uname -m)"
case "$uname_arch" in
  arm64|aarch64) arch="arm64" ;;
  x86_64|amd64) arch="x64" ;;
  *)
    echo "W365 CLI requires an x64 or arm64 machine. Detected: $uname_arch" >&2
    exit 1
    ;;
esac

if [ "$os" = "macos" ]; then
  asset_name="w365-osx-$arch.zip"
else
  asset_name="w365-linux-$arch.tar.gz"
fi

echo ""
echo "  W365 CLI"
echo "  --------"
echo "  OS           : $os"
echo "  Architecture : $arch"
echo ""

# --- Resolve latest release asset ------------------------------------------
echo "Looking up latest release..."
api_response="$(curl -fsSL \
  --retry 3 --retry-delay 2 \
  -H 'User-Agent: W365CLI-Installer' \
  -H 'Accept: application/vnd.github+json' \
  "https://api.github.com/repos/$REPO/releases/latest")" || {
    echo "Couldn't reach GitHub." >&2
    exit 1
  }

# GitHub returns this as one big single-line JSON blob (not pretty-printed), so line-anchored
# grep doesn't work here — use "grep -o" to pull out just the matching key/value substrings
# regardless of how the response is laid out.
release_tag="$(printf '%s' "$api_response" | grep -o '"tag_name":"[^"]*"' | head -1 | sed -E 's/^"tag_name":"//; s/"$//')"
download_url="$(printf '%s' "$api_response" | grep -o '"browser_download_url":"[^"]*"' | grep "$asset_name" | head -1 | sed -E 's/^"browser_download_url":"//; s/"$//')" || true

if [ -z "$download_url" ]; then
  available="$(printf '%s' "$api_response" | grep -o '"browser_download_url":"[^"]*"' | sed -E 's#.*/##; s/"$//' | tr '\n' ' ')"
  echo "Release ${release_tag:-latest} doesn't contain a W365 CLI build for $os/$arch. Available: $available" >&2
  exit 1
fi

# --- Download and extract ---------------------------------------------------
tmpdir="$(mktemp -d)"
trap 'rm -rf "$tmpdir"' EXIT

archive_path="$tmpdir/$asset_name"
echo "Downloading ${release_tag:-latest}..."
curl -fL --retry 5 --retry-delay 2 -o "$archive_path" "$download_url"

if [ ! -s "$archive_path" ]; then
  echo "Download failed or file is empty." >&2
  exit 1
fi

extract_dir="$tmpdir/extracted"
mkdir -p "$extract_dir"
if [ "$os" = "macos" ]; then
  unzip -q "$archive_path" -d "$extract_dir"
else
  # Linux releases ship as tar.gz (preserves the executable bit on extraction, unlike zip).
  tar -xzf "$archive_path" -C "$extract_dir"
fi

source_bin="$(find "$extract_dir" -type f -name 'W365Cli' -print -quit)"
if [ -z "$source_bin" ]; then
  echo "Couldn't find the W365Cli binary inside the downloaded archive." >&2
  exit 1
fi

# --- Install -----------------------------------------------------------------
mkdir -p "$INSTALL_DIR"
dest_bin="$INSTALL_DIR/$BIN_NAME"
cp "$source_bin" "$dest_bin"
chmod +x "$dest_bin"

if [ "$os" = "macos" ]; then
  # Defensive: if this file ever picks up a quarantine flag (e.g. downloaded via a browser
  # instead of curl), clear it. The binary is already signed and notarized in signed releases, so
  # this does not bypass Gatekeeper — it just avoids a redundant prompt for an already-trusted
  # binary. xattr doesn't exist on Linux, so this only ever runs on macOS.
  xattr -dr com.apple.quarantine "$dest_bin" 2>/dev/null || true
fi

echo ""
echo "W365 CLI installed to $dest_bin"

# --- PATH ---------------------------------------------------------------------
# Track whether the CURRENT shell's PATH (not the rc file) already includes INSTALL_DIR — these
# are two different things, and conflating them was a real bug: if a previous run of this script
# already added the export line to the rc file, but the current shell was never restarted since,
# grep would find the line "already there" and silently skip everything, while the shell running
# right now still couldn't find w365cli. That hits anyone who reinstalls/upgrades/re-runs this
# script without opening a fresh terminal in between — not just a one-off edge case.
path_updated=0
path_missing_from_shell=0
rc_file=""
rc_files=""
if [ "$NO_PATH" -eq 0 ]; then
  case ":$PATH:" in
    *":$INSTALL_DIR:"*)
      # Current shell's PATH already has it — nothing to do.
      ;;
    *)
      path_missing_from_shell=1
      shell_name="$(basename "${SHELL:-bash}")"
      case "$shell_name" in
        zsh) rc_files="$HOME/.zshrc" ;;
        bash)
          # macOS Terminal.app launches login shells (reads .bash_profile). Most Linux desktop
          # terminals launch non-login interactive shells (reads .bashrc) -- but WSL / Windows
          # Terminal launches LOGIN shells, and bash's login-shell startup reads ONLY the first
          # of .bash_profile / .bash_login / .profile it finds, completely skipping .bashrc (and
          # the other two) if that file exists. Some distro images (including stock Ubuntu-on-WSL
          # setups we've seen in the wild) ship with an empty ~/.bash_profile stub, which silently
          # shadows .profile and therefore .bashrc for every login shell -- so our .bashrc edit
          # never takes effect no matter how many new terminals you open. Always update .bashrc
          # (covers non-login shells) and ALSO update .bash_profile when it exists (covers login
          # shells wherever that stub is present), so the PATH change works either way.
          if [ "$os" = "macos" ]; then
            rc_file="$HOME/.bash_profile"
            rc_files="$rc_file"
          else
            rc_file="$HOME/.bashrc"
            rc_files="$rc_file"
            if [ -f "$HOME/.bash_profile" ]; then
              rc_files="$rc_files $HOME/.bash_profile"
            fi
          fi
          ;;
        *) rc_file="$HOME/.profile"; rc_files="$rc_file" ;;
      esac
      path_line='export PATH="$HOME/.local/bin:$PATH"'
      for f in $rc_files; do
        if [ -f "$f" ] && grep -qF "$path_line" "$f" 2>/dev/null; then
          : # Line already present from a prior run — the CURRENT shell just hasn't reloaded it yet.
        else
          printf '\n# Added by W365 CLI installer\n%s\n' "$path_line" >> "$f"
          path_updated=1
          echo "Added ~/.local/bin to your PATH in $f"
        fi
      done
      ;;
  esac
fi

echo ""
if [ "$NO_PATH" -eq 1 ]; then
  echo "Run it with: $dest_bin"
elif [ "$path_missing_from_shell" -eq 1 ]; then
  # Whether we just edited the rc file this run or it was already there from an earlier attempt,
  # this shell's PATH doesn't have it loaded yet either way — 'source' fixes it immediately in
  # this same window, no need to close and reopen the terminal.
  echo "Run 'source $rc_file' to use w365cli right now in this terminal (or just open a new one)."
else
  echo "Type 'w365cli' to get started."
fi
echo ""
