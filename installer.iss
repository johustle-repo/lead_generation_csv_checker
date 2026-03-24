#define MyAppName "Elmar's Lead Generation Quality Studio"
#define MyAppVersion "1.2.2"
#define MyAppPublisher "Elmar Noche"
#define MyAppExeName "ElmarsLeadGenerationQualityStudio.exe"
#define MyAppBuildDir "dist\\ElmarsLeadGenerationQualityStudio"
#define MyAppUninstallKey "{1C916B9D-7C3F-4B58-8E03-BC1B8C8C6A8C}_is1"
#define MyInstallerBaseName "ElmarsLeadGenerationQualityStudio"

[Setup]
AppId={{1C916B9D-7C3F-4B58-8E03-BC1B8C8C6A8C}
AppName={#MyAppName}
AppVerName={#MyAppName} {#MyAppVersion}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Elmar Lead Generation Quality Studio
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer-dist
OutputBaseFilename={#MyInstallerBaseName}-{#MyAppVersion}-Setup
SetupIconFile=app-logo.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Files]
Source: "{#MyAppBuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function TryGetPreviousUninstallCommand(var UninstallCommand: String): Boolean;
begin
  Result :=
    RegQueryStringValue(HKLM, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppUninstallKey}', 'QuietUninstallString', UninstallCommand) or
    RegQueryStringValue(HKLM, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppUninstallKey}', 'UninstallString', UninstallCommand) or
    RegQueryStringValue(HKLM64, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppUninstallKey}', 'QuietUninstallString', UninstallCommand) or
    RegQueryStringValue(HKLM64, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppUninstallKey}', 'UninstallString', UninstallCommand) or
    RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppUninstallKey}', 'QuietUninstallString', UninstallCommand) or
    RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppUninstallKey}', 'UninstallString', UninstallCommand);
end;

function EnsureSilentUninstallFlags(const UninstallCommand: String): String;
var
  UpperCommand: String;
begin
  Result := Trim(UninstallCommand);
  UpperCommand := UpperCase(Result);

  if Pos('/VERYSILENT', UpperCommand) = 0 then
    Result := Result + ' /VERYSILENT';
  if Pos('/SUPPRESSMSGBOXES', UpperCommand) = 0 then
    Result := Result + ' /SUPPRESSMSGBOXES';
  if Pos('/NORESTART', UpperCommand) = 0 then
    Result := Result + ' /NORESTART';
end;

function UninstallPreviousVersionIfPresent(): String;
var
  UninstallCommand: String;
  ResultCode: Integer;
begin
  Result := '';
  if not TryGetPreviousUninstallCommand(UninstallCommand) then
    exit;

  if not Exec(
    ExpandConstant('{cmd}'),
    '/C "' + EnsureSilentUninstallFlags(UninstallCommand) + '"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  ) then
  begin
    Result := 'Setup could not uninstall the previous version automatically. Please uninstall it first, then run this installer again.';
    exit;
  end;

  if ResultCode <> 0 then
    Result :=
      'The previous version uninstall did not complete successfully (exit code ' + IntToStr(ResultCode) +
      '). Please uninstall it first, then run this installer again.';
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
  ErrorMessage: String;
begin
  Exec(
    ExpandConstant('{cmd}'),
    '/C taskkill /IM "{#MyAppExeName}" /F /T >nul 2>nul',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  );

  ErrorMessage := UninstallPreviousVersionIfPresent();
  if ErrorMessage <> '' then
  begin
    Result := ErrorMessage;
    exit;
  end;

  Result := '';
end;
