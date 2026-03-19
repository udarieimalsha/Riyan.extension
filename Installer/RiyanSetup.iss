[Setup]
AppId={{5d17a437-4553-4a2e-80fa-cc207f37982c}
AppName=Riyan Revit Tools
AppVersion=1.0.1
AppPublisher=Riyan
AppPublisherURL=https://github.com/udarieimalsha/Riyan.extension
AppSupportURL=https://github.com/udarieimalsha/Riyan.extension
AppUpdatesURL=https://github.com/udarieimalsha/Riyan.extension
DefaultDirName={userappdata}\pyRevit\Extensions\Riyan.extension
DefaultGroupName=Riyan Revit Tools
AllowNoIcons=yes
OutputDir=..
OutputBaseFilename=RiyanSetup_v1.0.1
; SetupIconFile=..\Riyan.tab\About.panel\About.pushbutton\icon.png
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
WizardImageFile=WizardImage.png
WizardSmallImageFile=WizardSmallImage.png
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

; Version Info for Windows Properties
VersionInfoCompany=Riyan
VersionInfoDescription=Riyan Revit Tools Installer
VersionInfoVersion=1.0.1
VersionInfoCopyright=Copyright (C) 2026 Riyan

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Copy all extension files except the installer itself and git metadata
Source: "../*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; \
    Excludes: ".git\*,.gitignore,.gemini\*,Installer\*,*.md,*.exe,*.zip"

[Icons]
Name: "{group}\Uninstall Riyan Revit Tools"; Filename: "{uninstallexe}"

[Run]
; Optional: Reload pyRevit if the CLI is in the path
Filename: "pyrevit"; Parameters: "reload"; Flags: runhidden nowait skipifsilent; StatusMsg: "Reloading pyRevit..."

[InstallDelete]
; Clean up old files before installing new ones
Type: filesandordirs; Name: "{app}\*"
