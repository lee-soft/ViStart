; Inno Setup 6 script for ViStart (classic VB6 branch).
;
; Reconstructed 2026-08 after the 8.1.0.5311 script was lost. Footprint
; forensics came from a live 5311 install (unins000.dat + HKCU uninstall key):
; per-user install into {userappdata}\ViStart, payload = exe + Languages +
; Plugins + 13 skins into _skins. No orbs, no settings.xml (the app generates
; it on first run - shipping one would clobber user config on upgrade).
;
; AppId below reproduces the 5311 registry key EXACTLY, including the trailing
; "}}" that the original script's {{...}} brace-escaping produced. The HKCU
; uninstall key is "{D4837ABF-9483-4945-A490-2FE854E11259}}_is1" - keep the
; double closing brace or upgrades stop replacing the existing entry.
;
; ViStart.exe is compiled by ..\build.ps1, which outputs to %APPDATA%\ViStart.
; Override with: ISCC /DBuildDir=<dir> installer\ViStart.iss

#define MyAppName "ViStart"
; Marketing/PAD version format (matches the 2023 "8.1.0.5311" precedent; the
; exe's own FileVersion string is VB6-style "8.01.xxxx" - ViUpdater's version
; compare is numeric-tolerant so the two formats agree).
#define MyAppVersion "8.1.0.5332"
#define MyAppPublisher "Lee-Soft.com"
#define MyAppExeName "ViStart.exe"
#ifndef BuildDir
  #define BuildDir GetEnv("APPDATA") + "\ViStart"
#endif

[Setup]
AppId={{D4837ABF-9483-4945-A490-2FE854E11259}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://lee-soft.com/vistart/
AppSupportURL=https://lee-soft.com/vistart/
DefaultDirName={userappdata}\ViStart
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=Output
; /bin/ filename convention: version with the dots squashed out
OutputBaseFilename=windows-start-menu-vistart-{#StringChange(MyAppVersion, ".", "")}
SetupIconFile=..\ViStart.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoVersion={#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
; Built with Inno Setup 5.6.1 (Unicode) so the installer runs on Windows
; XP/2000 through 11 - "one exe, XP to 11" is part of the product's pitch.
; The script stays compatible with Inno 6; the guard just drops the
; 6-only wizard style when compiled with 5.x.
#if VER >= EncodeVer(6,0,0)
WizardStyle=modern
#endif
; We kill the running instance ourselves in PrepareToInstall; Restart Manager
; guesses wrong on a shell-replacement app.
CloseApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"

[Files]
Source: "{#BuildDir}\ViStart.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Languages\*"; DestDir: "{app}\Languages"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "..\Plugins\*"; DestDir: "{app}\Plugins"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "..\Skins\*"; DestDir: "{app}\_skins"; Flags: recursesubdirs createallsubdirs ignoreversion
; Default orb: the app falls back to _orbs\Windows 7.png when the configured
; orb and the skin's start_button.png are both missing, so ship it (5311's
; footprint had no orbs, which left a blank orb on skins without one).
; onlyifdoesntexist: never clobber an orb the user replaced.
Source: "..\Orbs\Windows 7.png"; DestDir: "{app}\_orbs"; Flags: onlyifdoesntexist

[Registry]
; Start with Windows. The app's control-panel checkbox reads/writes this same
; value and compares it UNQUOTED (Options.bas!StartsWithWindows), so the data
; must stay a bare path. The 5311 installer never wrote this - fresh installs
; silently didn't survive a reboot unless the user found the checkbox.
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ViStart"; ValueData: "{app}\{#MyAppExeName}"; Flags: uninsdeletevalue

[Icons]
Name: "{userprograms}\ViStart"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\ViStart"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "https://lee-soft.com/vistart/thankyou/"; Flags: shellexec nowait skipifsilent
Filename: "{app}\{#MyAppExeName}"; Description: "Launch ViStart"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM ViStart.exe /IM MetroProvider.exe /IM SearchProvider.exe"; Flags: runhidden skipifdoesntexist; RunOnceId: "KillViStart"

[Code]
// Stop the running instance (and plugin processes) so their files can be
// replaced. Plain /F by image name, never /T: when ViUpdater launches this
// installer a tree-kill could take out our own process chain.
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'),
    '/F /IM ViStart.exe /IM MetroProvider.exe /IM SearchProvider.exe',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(500);
  Result := '';
end;

// Machines upgraded from the NSIS era (<= 8.1.0.5208) carry a stale
// HKCU\...\Uninstall\ViStart key whose KillMe.exe uninstaller no longer
// exists, leaving a dead "ViStart" row in Add/Remove next to ours. Remove it
// only when its uninstaller is genuinely gone.
procedure CurStepChanged(CurStep: TSetupStep);
var
  UninstallString: String;
begin
  if CurStep = ssPostInstall then
    if RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\ViStart',
        'UninstallString', UninstallString) then
      if not FileExists(RemoveQuotes(UninstallString)) then
        RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\ViStart');
end;
