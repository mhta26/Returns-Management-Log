#define AppVersion "1.0.0"
[Setup]
AppId={{C0AEE46B-20BA-4051-AC73-0D4179E11AA4}
AppName=Shanifi Returns Manager
AppVersion={#AppVersion}
DefaultDirName={autopf}\Shanifi Returns Manager
DefaultGroupName=Shanifi Returns Manager
OutputDir=..\release
OutputBaseFilename=ReturnsManager-Setup-Windows-x64
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern
UninstallDisplayIcon={app}\Client\ReturnsManager.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
Source: "..\dist\ReturnsManager\*"; DestDir: "{app}\Client"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\dist\ReturnsServer\*"; DestDir: "{app}\Server"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\OPERATIONS.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\THIRD_PARTY.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Returns Manager"; Filename: "{app}\Client\ReturnsManager.exe"
Name: "{group}\Start Shared Server"; Filename: "{app}\Server\ReturnsServer.exe"
Name: "{autodesktop}\Returns Manager"; Filename: "{app}\Client\ReturnsManager.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Client\ReturnsManager.exe"; Description: "Launch Returns Manager"; Flags: nowait postinstall skipifsilent
