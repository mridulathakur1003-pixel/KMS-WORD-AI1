#define MyAppName "KMS WORD AI V20 PRO"
#define MyAppVersion "20.0.1"
#define MyAppPublisher "KMS A VISION"
#define MyAppExeName "KMS WORD AI.exe"

[Setup]
AppId={{B3C21E62-18E2-4D8D-8A6C-2F7B1A0C9020}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\KMS WORD AI
DefaultGroupName=KMS WORD AI
OutputDir=..\release
OutputBaseFilename=KMS_WORD_AI_V20_PRO_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\dist\KMS WORD AI\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\KMS WORD AI V20 PRO"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\KMS WORD AI V20 PRO"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch KMS WORD AI"; Flags: postinstall nowait skipifsilent
