# Security Policy

## Reporting a Vulnerability

If you discover a security vulnerability in W365 CLI, please **do not open a public GitHub issue**.
Instead, report it privately using one of these methods:

1. **GitHub Security Advisories (preferred)** — open a
   [private security advisory](https://github.com/bwya77/W365-CLI-Native/security/advisories/new)
   for this repository. This lets us coordinate a fix and disclosure timeline without exposing the
   issue publicly before a patch is available.
2. **Email** — if you can't use GitHub Security Advisories, contact the maintainer directly (see
   the GitHub profile at [bwya77](https://github.com/bwya77) for contact details).

Please include as much detail as you can:

- A description of the vulnerability and its potential impact
- Steps to reproduce, including affected version(s)
- Any proof-of-concept code or captured network traffic (redact tenant-specific identifiers)

You should receive an acknowledgment within a few days. We'll keep you updated as the issue is
triaged, fixed, and released, and we're happy to credit reporters in the release notes unless you
prefer to stay anonymous.

## Supported Versions

Only the latest released version of W365 CLI is supported with security fixes. Because the CLI
checks for updates on startup and can update itself in place (see the
[Updates](README.md#updates) section of the README), users are expected to stay on the latest
release.

## What This Tool Can Access

W365 CLI is a Microsoft Graph client. Understanding what it can do helps you evaluate its security
posture:

- It authenticates interactively via MSAL using **delegated permissions** — it only acts with the
  privileges of the signed-in user/admin, never with its own standing access to your tenant.
- It uses a **public client app registration** (`9d497858-c200-402c-a363-279a5800d730`) with a
  `http://localhost` redirect URI and no client secret — there is nothing secret embedded in the
  binary that, if extracted, would grant an attacker tenant access on its own.
- The delegated Graph permissions it requests are listed in full in the
  [Permissions](README.md#permissions) section of the README, including which ones are strictly
  required vs. optional (e.g. `GroupMember.ReadWrite.All` is only needed for the group-membership
  management feature).
- Authentication tokens are cached locally via MSAL's standard token cache, protected by the OS
  credential store where available (Windows DPAPI / Credential Manager, macOS Keychain, Linux
  Secret Service via libsecret) and falling back to a user-only-readable file when no OS keyring is
  available (e.g. headless Linux). W365 CLI never transmits or logs your tokens.
- Every Microsoft Graph call the CLI makes goes directly from your machine to
  `graph.microsoft.com` (or `login.microsoftonline.com` for auth) over TLS — there is no
  intermediary telemetry, logging, or analytics backend operated by this project.
- The CLI does not collect telemetry, analytics, or usage data of any kind.

## Supply Chain / Build Integrity

- **Source is fully public** in this repository — nothing in the release builds is built from code
  you can't read yourself.
- **Windows** release binaries and installers are signed with **Azure Trusted Signing**.
- **macOS** release binaries are signed and notarized with Apple when the maintainer's Developer
  ID signing secrets are available for that release.
- **Linux** release binaries are not code-signed (no equivalent ecosystem convention exists for
  Linux CLI binaries); verify integrity against the published `SHA256SUMS-linux.txt` checksum file.
- Every release publishes `SHA256SUMS-*.txt` files alongside the binaries — always verify a
  downloaded asset's checksum before running it if you have any doubt about its provenance.
- Releases are built entirely by GitHub Actions from this repository's own workflow files
  (`.github/workflows/release.yml`), not built or uploaded manually — you can read the exact build
  steps that produced any given release.
- Dependencies are limited to a small, well-known set (Microsoft.Identity.Client/MSAL,
  Spectre.Console, System.Text.Json) and are kept current via Dependabot.

## Reporting Non-Security Issues

For bugs that aren't security-sensitive, please use the normal
[GitHub Issues](https://github.com/bwya77/W365-CLI-Native/issues) tracker instead.
