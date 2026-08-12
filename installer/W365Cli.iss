; W365 CLI installer (Inno Setup). Built in CI, once per architecture:
;
;   ISCC /DAppVersion=<version> /DArch=x64|arm64 /DSourceDir=<publish dir> /O<out dir> installer\W365Cli.iss
;
; Installs the self-contained CLI into a per-user Program Files-style folder (no admin/UAC
; required), optionally adds that folder to the user's PATH so "w365cli" works from any new
; terminal, creates a Start Menu shortcut, and registers a normal Windows uninstaller (shows up
; under Settings > Apps / Control Panel > Programs and Features, just like any other app).

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef Arch
  #define Arch "x64"
#endif
#ifndef SourceDir
  #define SourceDir "..\publish\win-" + Arch
#endif

[Setup]
AppId={{7BA17D73-6A35-4892-82CA-CA4AB563BAAD}
AppName=W365 CLI
AppVersion={#AppVersion}
AppPublisher=Bradley Wyatt
AppPublisherURL=https://github.com/bwya77/W365-CLI-Native
AppSupportURL=https://github.com/bwya77/W365-CLI-Native/issues
AppUpdatesURL=https://github.com/bwya77/W365-CLI-Native/releases
; Per-user install under %LocalAppData%\Programs — no admin/UAC prompt required, matching modern
; CLI tool installer conventions (gh, deno, uv, etc.). Still registers a normal uninstall entry.
DefaultDirName={localappdata}\Programs\W365CLI
DefaultGroupName=W365 CLI
DisableProgramGroupPage=yes
DisableDirPage=auto
UsePreviousAppDir=no
UninstallDisplayIcon={app}\W365Cli.exe
UninstallDisplayName=W365 CLI
OutputBaseFilename=W365CLISetup-{#AppVersion}-win-{#Arch}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; No admin rights needed for a per-user install/PATH change — avoids a UAC prompt entirely.
PrivilegesRequired=lowest
; Close a running W365 CLI before replacing files so an in-place update never fails on a locked exe.
CloseApplications=yes
RestartApplications=no
#if Arch == "arm64"
ArchitecturesAllowed=arm64
ArchitecturesInstallIn64BitMode=arm64
#else
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
#endif

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Excludes: "*.pdb"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\W365 CLI"; Filename: "{app}\W365Cli.exe"

[Tasks]
Name: "addtopath"; Description: "Add W365 CLI to PATH (lets you run ""w365cli"" from any terminal)"; GroupDescription: "Additional options:"; Flags: checkedonce

[Code]
const
  EnvironmentKey = 'Environment';

// Belt-and-suspenders process kill right before file copy — CloseApplications above uses
// Windows Restart Manager, which is usually enough, but this guarantees an in-place update
// never fails because a terminal still has the exe open.
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM W365Cli.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(200);
  Result := '';
end;

procedure AddToPath();
var
  OrigPath: string;
  AppDir: string;
begin
  AppDir := ExpandConstant('{app}');
  if not RegQueryStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath) then
    OrigPath := '';
  if Pos(';' + AppDir + ';', ';' + OrigPath + ';') = 0 then
  begin
    if (Length(OrigPath) > 0) and (OrigPath[Length(OrigPath)] <> ';') then
      OrigPath := OrigPath + ';';
    OrigPath := OrigPath + AppDir;
    RegWriteStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath);
  end;
end;

procedure RemoveFromPath();
var
  OrigPath, NewPath: string;
  P: Integer;
  AppDir: string;
begin
  AppDir := ExpandConstant('{app}');
  if not RegQueryStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath) then
    exit;
  NewPath := ';' + OrigPath + ';';
  P := Pos(';' + AppDir + ';', NewPath);
  if P > 0 then
  begin
    Delete(NewPath, P, Length(AppDir) + 1);
    if (Length(NewPath) > 0) and (NewPath[1] = ';') then
      Delete(NewPath, 1, 1);
    if (Length(NewPath) > 0) and (NewPath[Length(NewPath)] = ';') then
      Delete(NewPath, Length(NewPath), 1);
    RegWriteStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', NewPath);
  end;
end;

// Broadcast WM_SETTINGCHANGE so processes that pick up environment variable changes live
// (e.g. Explorer) notice the PATH update without a full sign-out/sign-in. Terminals that were
// already open before install/uninstall still need to be reopened to see the new PATH — this
// is a Windows limitation, not something an installer can force onto other running processes.
function SendMessageTimeoutA(hWnd: LongInt; Msg: LongInt; wParam: LongInt; lParam: string;
  fuFlags: LongInt; uTimeout: LongInt; var lpdwResult: LongInt): LongInt;
  external 'SendMessageTimeoutA@user32.dll stdcall';

procedure RefreshEnvironment();
var
  ResultCode: LongInt;
begin
  SendMessageTimeoutA($FFFF, $001A, 0, 'Environment', 2, 5000, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssPostInstall) and WizardIsTaskSelected('addtopath') then
  begin
    AddToPath();
    RefreshEnvironment();
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveFromPath();
    RefreshEnvironment();
  end;
end;
