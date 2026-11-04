; =====================================================================
; Inno Setup Script - Civil Tool (MyFirstProject)
; Tac Gia: Huynh Ngoc Nam
; =====================================================================

#define AppName      "Bo Cong Cu Civil 3D - NXSOFT"
#ifndef AppVersion
  #define AppVersion "1.2.0"
#endif
#define AppPublisher "Huynh Ngoc Nam"
#define AppURL       "Namhuynh1205@gmail.com"
#define SourceDir    "..\MyFirstProject\bin\Release"
#define OutputDir    "..\..\Civil_Tool_Installer"
#define BundleName   "MyFirstProject.bundle"

[Setup]
AppId={{DA4B24B3-D7F4-44C9-8802-A62E8F18579C}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}

; Cai dat thang vao thu muc Plugin cua AutoCAD (ApplicationPlugins)
DefaultDirName={commonappdata}\Autodesk\ApplicationPlugins\{#BundleName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
AllowNoIcons=yes

OutputDir={#OutputDir}
OutputBaseFilename=Civil_Tool_NXSOFT_Setup_v{#AppVersion}
; SetupIconFile: file logo.ico cu (FolderTreeGenerator) khong con ton tai -> tam bo
; SetupIconFile=D:\Autocad 2026_API\FolderTreeGenerator\logo.ico

Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern

; Can quyen admin do cai vao ProgramData
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
english.WelcomeLabel1=Chao mung ban den voi trinh cai dat [bold]{#AppName}[/bold]
english.WelcomeLabel2=Phan mem nay se tu dong tich hop bo cong cu vao AutoCAD Civil 3D tren may tinh cua ban.%n%nKhuyen nghi tat AutoCAD truoc khi tien hanh cai dat.
english.FinishedLabel=Cai dat [name] da hoan tat. Tu nay Plugin se tu dong load khi ban mo AutoCAD.

[InstallDelete]
Type: filesandordirs; Name: "{app}\Contents"

[Files]
; Copy file PackageContents vao goc cua thu muc .bundle
Source: "PackageContents.xml"; DestDir: "{app}"; Flags: ignoreversion

; Copy toan bo file cua plugin vao thu muc Contents, THEO QUY TAC LOAI BO cac dll cua AutoCAD (de tranh loi sung dot dll)
Source: "{#SourceDir}\*"; DestDir: "{app}\Contents"; Excludes: "accoremgd.dll,acdbmgd.dll,acmgd.dll,acdbmgdbrep.dll,AcCui.dll,AecBaseMgd.dll,AeccDbMgd.dll,AdWindows.dll,AecPropDataMgd.dll,Autodesk.AutoCAD.Interop.dll,Autodesk.AutoCAD.Interop.Common.dll,AecBaseMgd.xml,AecPropDataMgd.xml,Ac32BitAppServer.dll,AcCustomize.dll,AcDx.dll,AcDxPublishUi.dll,AcJsonSerializable.dll,AcMNUParser.dll,AcMr.dll,AcTcMgd.dll,AcWindows.dll,AdUIMgd.dll,AdUiPalettes.dll,Aecc*.dll,DesktopConnectorInterop.dll,System.ServiceModel*.dll,Microsoft.CodeAnalysis*.dll,*.deps.json"; Flags: ignoreversion recursesubdirs createallsubdirs

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
