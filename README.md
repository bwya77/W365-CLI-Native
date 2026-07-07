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
- An Entra public client app registration for this native CLI.
- Microsoft Graph delegated permissions needed for Windows 365 scenarios.

Set the app registration client ID before running:

```powershell
$env:W365CLI_CLIENT_ID = '<client-id>'
```

Optionally set a tenant ID:

```powershell
$env:W365CLI_TENANT_ID = '<tenant-id>'
```

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
