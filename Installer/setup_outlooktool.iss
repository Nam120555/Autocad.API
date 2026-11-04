; =====================================================================
; Inno Setup Script - OutLook Email Filter & Export Tool
; Tac Gia: Huynh Ngoc Nam - NXSOFT
; =====================================================================

#define AppName      "Bo Cong Cu Loc Email Outlook - NXSOFT"
#define AppVersion   "1.0.0"
#define AppPublisher "Huynh Ngoc Nam"
#define AppExeName   "OutLookTool.exe"
#define SourceDir    "C:\OutLook"
#define OutputDir    "..\OutLookTool_Installer"

[Setup]
AppId={{C8E110A5-492B-4813-A3B9-8F7BD5C82A41}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisher={#AppPublisher}

DefaultDirName=C:\OutLook
DefaultGroupName={#AppName}
AllowNoIcons=yes

OutputDir={#OutputDir}
OutputBaseFilename=Setup_OutLookTool_v{#AppVersion}
SetupIconFile=D:\Autocad 2026_API\FolderTreeGenerator\logo.ico

Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#SourceDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\*.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#SourceDir}\Chay_Tool_Doc_Email_OutLook.bat"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "D:\Autocad 2026_API\FolderTreeGenerator\logo.ico"; DestDir: "{app}"; DestName: "app_icon.ico"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\app_icon.ico"
Name: "{group}\Go Cai Dat {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\app_icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Chay {#AppName} ngay bay gio"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\inject_vba_macro.ps1"
Type: dirifempty; Name: "{app}\_TempRead"
