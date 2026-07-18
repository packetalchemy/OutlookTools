; OutlookTools Inno Setup Script v1.2.0
; Compile with Inno Setup 6+

#define MyAppName "OutlookTools"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "OutlookTools Contributors"
#define MyAppURL "https://github.com/packetalchemy/OutlookTools"

[Setup]
AppId={{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=Output
OutputBaseFilename=OutlookTools_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
MinVersion=10.0.17763

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\OutlookTools\bin\Release\OutlookTools.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\OutlookTools\bin\Release\OutlookTools.pdb"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\docs\Documentation.html"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\docs\Documentation.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName} Documentation"; Filename: "{app}\docs\Documentation.html"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Registry]
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: string; ValueName: "Description"; ValueData: "OutlookTools"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OutlookTools"; Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Flags: uninsdeletekey

[Run]
Filename: "{sys}\regasm.exe"; Parameters: "/codebase /tlb ""{app}\OutlookTools.dll"""; StatusMsg: "Registering..."; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "{sys}\regasm.exe"; Parameters: "/unregister ""{app}\OutlookTools.dll"""; Flags: runhidden waituntilterminated

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
