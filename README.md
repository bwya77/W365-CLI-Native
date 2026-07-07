# W365 CLI Native

Native .NET rewrite experiment for Windows 365 Cloud PC workflows.

This repo is separate from the PowerShell-based `W365CLI` module. The goal is to learn whether a fully native binary can provide the same workflows without depending on the PowerShell module runtime.

## Current status

This is an early scaffold. It currently includes:

- Spectre.Console menu shell.
- Interactive browser authentication using Azure.Identity.
- Initial Cloud PC inventory table.
- Initial Cloud Apps table.

## Prerequisites

- .NET 8 SDK.
- Microsoft Graph delegated permissions consented for the W365 CLI Native app.

The default app registration client ID is built in:

```text
9d497858-c200-402c-a363-279a5800d730
```

The app registration must be configured as a native public client:

1. Open the app registration in Microsoft Entra.
2. Go to **Authentication**.
3. Add a **Mobile and desktop applications** platform.
4. Add this redirect URI:

```text
http://localhost
```

5. In **Advanced settings**, set **Allow public client flows** to **Yes**.

You can override it during development:

```powershell
$env:W365CLI_CLIENT_ID = '<client-id>'
```

Optionally set a tenant ID:

```powershell
$env:W365CLI_TENANT_ID = '<tenant-id>'
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

The app uses a persistent token cache so users do not need to sign in every run when refresh tokens are available.

## Build

```powershell
dotnet build
```

## Run

```powershell
dotnet run --project .\src\W365Cli\W365Cli.csproj
```

## Publish a local binary

Framework-dependent:

```powershell
dotnet publish .\src\W365Cli\W365Cli.csproj -c Release -o .\artifacts\publish
```

Single-file Windows x64:

```powershell
dotnet publish .\src\W365Cli\W365Cli.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o .\artifacts\publish\win-x64
```

## Direction

The native rewrite should eventually cover:

- Connection management.
- Cloud PC inventory and lifecycle actions.
- Snapshots.
- Provisioning policies.
- Maintenance windows.
- Cloud Apps publish and unpublish.
- Reports and usage views.

The first milestone is read-only browsing with a reliable Graph client and terminal UI.
