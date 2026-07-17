; OutlookTools Inno Setup Script
; Run this with Inno Setup Compiler to create Setup.exe

#define MyAppName "OutlookTools"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "OutlookTools Contributors"
#define MyAppURL "https://github.com/packetalchemy/OutlookTools"
#define MyAppExeName "OutlookTools.dll"

[Setup]
AppId={{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=OutlookTools_Setup_{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=..\Resources\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Setup
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
MinVersion=10.0.17763

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "farsi"; MessagesFile: "compiler:Languages\Farsi.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main DLL
Source: "..\bin\Release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Required files
Source: "..\bin\Release\OutlookTools.pdb"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\bin\Release\OutlookTools.xml"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

; Documentation
Source: "..\docs\Documentation.html"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\docs\Documentation.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName} Documentation"; Filename: "{app}\docs\Documentation.html"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Registry]
; Register COM Add-in for Outlook
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: string; ValueName: "Description"; ValueData: "OutlookTools - Open-source Outlook Add-in"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OutlookTools"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: string; ValueName: "Manifest"; ValueData: "{app}\{#MyAppExeName}"; Flags: uninsdeletekey

[Run]
; Register COM DLL after installation
Filename: "{sys}\regasm.exe"; Parameters: "/codebase /tlb ""{app}\{#MyAppExeName}"""; StatusMsg: "Registering OutlookTools..."; Flags: runhidden waituntilterminated

; Unregister before uninstall
[UninstallRun]
Filename: "{sys}\regasm.exe"; Parameters: "/unregister ""{app}\{#MyAppExeName}"""; Flags: runhidden waituntilterminated

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{localappdata}\OutlookTools"

[Code]
// Check if Outlook is running
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  // Check if Outlook is running
  if RegKeyExists(HKCU, 'Software\Microsoft\Office\16.0\Outlook') or
     RegKeyExists(HKCU, 'Software\Microsoft\Office\15.0\Outlook') then
  begin
    if MsgBox('Outlook is currently running. It will be closed to complete installation. Continue?',
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      Exec('taskkill', '/f /im outlook.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      Sleep(2000);
    end
    else
      Result := False;
  end;
end;

// Check if .NET 4.8 is installed
function IsDotNet48Installed(): Boolean;
var
  ResultCode: Integer;
begin
  Result := RegKeyExists(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Release');
end;

function InitializeUninstall(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  // Close Outlook before uninstall
  Exec('taskkill', '/f /im outlook.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(1000);
end;
