# W365 CLI Native

[![CI](https://github.com/bwya77/W365-CLI-Native/actions/workflows/ci.yml/badge.svg)](https://github.com/bwya77/W365-CLI-Native/actions/workflows/ci.yml)
[![Release](https://github.com/bwya77/W365-CLI-Native/actions/workflows/release.yml/badge.svg)](https://github.com/bwya77/W365-CLI-Native/actions/workflows/release.yml)
[![Latest release](https://img.shields.io/github/v/release/bwya77/W365-CLI-Native?label=release)](https://github.com/bwya77/W365-CLI-Native/releases/latest)
[![.NET 8](https://img.shields.io/badge/.NET-8.0-512BD4)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Platform](https://img.shields.io/badge/platform-Windows-4091f2)](https://github.com/bwya77/W365-CLI-Native/releases)

W365 CLI Native is a keyboard-first Windows 365 Cloud PC management experience built as a native
.NET command-line app.

It is separate from the PowerShell-based `W365CLI` module and does not require the PowerShell module at runtime.

## What it does

- Browse, filter, sort, and inspect Cloud PCs.
- Run Cloud PC actions such as sync, restart, resize, rename, reprovision, power on,
  reset local admin password, and end grace period.
- View disk space, snapshots, and remote action history.
- Create, restore, and delete snapshots.
- Browse provisioning policies and view the Cloud PCs assigned to a policy.
- Export, copy, reprovision, and delete provisioning policies.
- Browse reports for usage, connectivity history, launch details, and Graph report streams.
- Browse Cloud Apps and publish or unpublish them.
- Browse service plans, gallery images, custom images, and supported regions.
- View tenant settings, setting profiles, and user settings.
- Check GitHub Releases for newer builds.

## Install

Download the latest release:

```text
https://github.com/bwya77/W365-CLI-Native/releases/latest
```

Download `w365-win-x64.zip`, extract it, and run `W365Cli.exe`.

Recommended install folder:

```text
%LOCALAPPDATA%\Programs\W365CLI
```

Add that folder to your user PATH if you want to launch the CLI from any terminal.

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

The app registration must be configured as a native public client with this redirect URI:

```text
http://localhost
```

Recommended delegated Microsoft Graph permissions:

```text
CloudPC.Read.All
CloudPC.ReadWrite.All
DeviceManagementManagedDevices.Read.All
DeviceManagementManagedDevices.PrivilegedOperations.All
User.Read.All
Group.Read.All
offline_access
openid
profile
email
```

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
- Delete a policy

### Reports

Reports include:

- Usage
- Connectivity history
- Launch details
- Cloud PC report streams

Where possible, selecting a Cloud PC row opens that Cloud PC's detail page.

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

Use **Check for updates** in the app to compare your local binary with the latest GitHub Release.

Release builds are published as GitHub Release assets:

```text
w365-win-x64.zip
w365-win-x64.zip.sha256
```

Release binaries are signed before packaging.

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

Create a release:

```powershell
git tag v0.1.0
git push origin v0.1.0
```
