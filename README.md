# W365 CLI

[![CI](https://github.com/bwya77/W365-CLI-Native/actions/workflows/ci.yml/badge.svg)](https://github.com/bwya77/W365-CLI-Native/actions/workflows/ci.yml)
[![Release](https://github.com/bwya77/W365-CLI-Native/actions/workflows/release.yml/badge.svg)](https://github.com/bwya77/W365-CLI-Native/actions/workflows/release.yml)
[![Latest release](https://img.shields.io/github/v/release/bwya77/W365-CLI-Native?label=release)](https://github.com/bwya77/W365-CLI-Native/releases/latest)
[![.NET 8](https://img.shields.io/badge/.NET-8.0-512BD4)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-4091f2)](https://github.com/bwya77/W365-CLI-Native/releases)

<p align="center">
  <img src="docs/images/MainUI.png" width="820" alt="W365 CLI main menu">
</p>

<details>
<summary><strong>More screenshots</strong></summary>
<br>

| | |
|---|---|
| ![Cloud PCs](docs/images/CloudPCs.png) | ![Disk space](docs/images/CloudPCDiskSpace.png) |
| Browse and filter your Cloud PC fleet | Inspect disk space and snapshots |
| ![Provisioning policies](docs/images/ProvisioningPolicies.png) | ![Policy Cloud PCs](docs/images/ProvisioningPoliciesCloudPCs.png) |
| Browse provisioning policies | View the Cloud PCs assigned to a policy |
| ![Cloud Apps](docs/images/CloudApps.png) | ![User experience sync](docs/images/UserExperienceSync.png) |
| Browse and publish Cloud Apps | Manage user experience sync storage and profiles |

</details>


W365 CLI is a keyboard-first Windows 365 Cloud PC management experience built as a .NET
command-line app.

It is separate from the PowerShell-based `W365CLI` module and does not require the PowerShell module at runtime.

## What it does

- Browse, filter, sort, and inspect Cloud PCs.
- Run Cloud PC actions such as sync, restart, resize, rename, reprovision, power on,
  reset local admin password, and end grace period.
- View disk space, snapshots, and remote action history.
- Create, restore, and delete snapshots.
- View detailed status information (including provisioning warnings/errors) for a Cloud PC.
- Browse provisioning policies and view the Cloud PCs assigned to a policy.
- Create a new provisioning policy from scratch.
- Export, copy, reprovision, and delete provisioning policies.
- Reprovision a shared policy's Cloud PCs while keeping a reserve percentage available.
- View and manage user experience sync (user settings persistence) storage and profiles for
  shared provisioning policies, including a tenant-wide overview across all policies.
- View and manage the members of a provisioning policy's assigned Entra group.
- Understand Windows 365 license capacity, availability, and Flex utilization.
- Browse reports for usage, connectivity history, launch details, and Graph report streams.
- Browse Cloud Apps and publish or unpublish them.
- Browse service plans, gallery images, custom images, and supported regions.
- View tenant settings, setting profiles, and user settings.
- Check GitHub Releases for newer builds and install updates automatically.

## Install

### Windows (recommended: installer)

One-line install (no admin/UAC prompt — installs to `%LocalAppData%\Programs\W365CLI` and adds it
to your user PATH):

```powershell
irm https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/install.ps1 | iex
```

Open a new terminal and type `w365cli` to get started. The installer also registers a normal
uninstaller under **Settings > Apps** (or **Control Panel > Programs and Features**), so you can
remove it like any other application — or run:

```powershell
irm https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/uninstall.ps1 | iex
```

Alternatively, download `W365CLISetup-<version>-win-x64.exe` or `-win-arm64.exe` from the
[latest release](https://github.com/bwya77/W365-CLI-Native/releases/latest) and run it directly.

### macOS (recommended: install script)

One-line install (no `sudo` required — installs to `~/.local/bin/w365cli` and adds it to your
PATH):

```bash
curl -fsSL https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/install.sh | bash
```

Open a new terminal and type `w365cli` to get started. To uninstall:

```bash
curl -fsSL https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/uninstall.sh | bash
```

Add `--purge-path` to also remove the PATH line the installer added to your shell profile
(`~/.zshrc` or `~/.bash_profile`).

### Portable (no install)

Download the latest release:

```text
https://github.com/bwya77/W365-CLI-Native/releases/latest
```

Download the package for your platform, extract it, and run the binary:

| Platform | Asset | Binary |
| --- | --- | --- |
| Windows x64 | `w365-win-x64.zip` | `W365Cli.exe` |
| Windows ARM64 | `w365-win-arm64.zip` | `W365Cli.exe` |
| macOS Intel | `w365-osx-x64.zip` | `W365Cli` |
| macOS Apple Silicon | `w365-osx-arm64.zip` | `W365Cli` |

On macOS, the MSAL token cache is stored with Keychain protection.

If you extract a portable package instead of using an installer/install script, add the
extracted folder to your PATH manually if you want to launch the CLI from any terminal.

## Sign in

The CLI uses Microsoft Graph delegated permissions and an interactive browser sign-in. After the
first successful sign-in, MSAL keeps a persistent token cache so you usually do not need to sign in
every run.

On startup, the CLI tries to silently restore the cached Microsoft Graph session. If no cached session exists, open:

```text
Connection > Connect
```

## App registration

The default public client app ID is built in:

```text
9d497858-c200-402c-a363-279a5800d730
```

The app registration must be configured as a public client (the "Mobile and desktop
applications" platform in Entra) with this redirect URI:

```text
http://localhost
```

Recommended delegated Microsoft Graph permissions:

```text
CloudPC.ReadWrite.All
DeviceManagementManagedDevices.Read.All
DeviceManagementManagedDevices.PrivilegedOperations.All
Group.Read.All
GroupMember.ReadWrite.All
User.Read.All
Organization.Read.All
offline_access
openid
profile
email
```

`GroupMember.ReadWrite.All` is required for the "Manage group members" feature (add/remove
members of a provisioning policy's assigned Entra group). `User.Read.All` is used to search the
directory when adding a member. If you only need read-only features, `Group.Read.All` alone is
enough and `GroupMember.ReadWrite.All` can be omitted.

You can override the client or tenant during development:

```powershell
$env:W365CLI_CLIENT_ID = '<client-id>'
$env:W365CLI_TENANT_ID = '<tenant-id>'
```

## Navigation

The CLI is designed for keyboard use.

| Key | Action |
| --- | --- |
| `Up` / `Down` | Move selection |
| `PgUp` / `PgDn` | Page through long tables |
| `Enter` | Open the selected row or run the selected action |
| `/` or `F` | Filter a table |
| `C` | Clear the current filter |
| `S` | Cycle table sort modes where available |
| `R` | Refresh data where available |
| `Esc`, `B`, or `Q` | Go back |
| `P` or `Ctrl+K` | Open command palette |
| `H` | Open in-session action history |

Action submissions show a brief result screen, then return to the previous page.

## Main areas

### Cloud PCs

The Cloud PCs area includes:

- Browse Cloud PCs
- Disk space across all Cloud PCs
- Snapshots across all Cloud PCs

Selecting a Cloud PC opens its detail page with actions and subviews for disk space, snapshots,
resize, and remote action history.

### Provisioning

The Provisioning area includes a provisioning policy browser with actions to:

- View Cloud PCs assigned to a policy
- Export policy JSON
- Create a policy copy
- Reprovision Cloud PCs assigned to the policy
- Reprovision a shared policy while keeping a reserve percentage available, and check status
- View user experience sync storage usage and profiles for shared-by-Entra-group policies
- Manage the members of a policy's assigned Entra group (view, add, remove)
- Delete a policy

It also includes a "Create policy" wizard, and a tenant-wide "User experience sync overview" that
rolls up storage usage across every shared-by-Entra-group policy.

### Reports

Reports include:

- Sign-in status
- Usage
- Connectivity history
- Launch details
- Cloud PC report streams

Where possible, selecting a Cloud PC row opens that Cloud PC's detail page.

### Licensing

Licensing summarizes Windows 365 subscribed SKU capacity against the current Cloud PC inventory.
It shows purchased licenses, assigned licenses, provisioned Cloud PCs, estimated availability,
Reserve Cloud PC usage, and Windows 365 Flex dedicated/shared utilization.

For Flex, the CLI shows the 3:1 dedicated provisioning model and the active-session limit. It also
shows provisioning policies and groups that grant access where Graph exposes assignment data.

### Catalog

Catalog includes:

- Service plans
- Gallery images
- Custom images
- Supported regions

### Tenant settings

Tenant settings includes:

- Organization settings
- Setting profiles
- User settings

### Cloud Apps

Cloud Apps includes browse, publish, and unpublish workflows.

## Updates

W365 CLI checks GitHub Releases on startup and offers to update automatically when a newer
version is available:

- **Windows** — downloads the matching installer and, if you agree, runs it silently
  (`/VERYSILENT /NORESTART`); W365 CLI closes for a few seconds while it updates, then reopen it.
  If you'd rather update later, the installer is saved locally and you can double-click it anytime.
- **macOS** — downloads the matching build and replaces the installed `w365cli` binary in place
  (an atomic rename, safe even while it's running), then offers to restart into the new version.

Release builds are published as GitHub Release assets:

```text
W365CLISetup-<version>-win-x64.exe
W365CLISetup-<version>-win-arm64.exe
w365-win-x64.zip
w365-win-arm64.zip
w365-osx-x64.zip
w365-osx-arm64.zip
SHA256SUMS-windows.txt
SHA256SUMS-macos.txt
```

Windows release binaries and installers are signed with Azure Trusted Signing before packaging.
macOS release binaries are signed and notarized when the Apple Developer signing secrets are
available.

## Development

Prerequisite:

```text
.NET 8 SDK
```

Build:

```powershell
dotnet build --configuration Release
```

Run:

```powershell
dotnet run --project .\src\W365Cli\W365Cli.csproj
```

Publish a local self-contained Windows x64 binary:

```powershell
dotnet publish .\src\W365Cli\W365Cli.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false -o .\artifacts\publish\win-x64
```

Other release runtime identifiers:

```text
win-arm64
osx-x64
osx-arm64
```

Create a release:

```powershell
git tag v0.2.0
git push origin v0.2.0
```
