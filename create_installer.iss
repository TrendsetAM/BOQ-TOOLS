
[Setup]
AppName=BOQ Tools
AppVersion=1.0.0
AppPublisher=Your Company
AppPublisherURL=https://your-website.com
DefaultDirName={pf}\BOQ Tools
DefaultGroupName=BOQ Tools
OutputDir=C:\Users\Admin\Documents\GitHub\BOQ-TOOLS\installer
OutputBaseFilename=BOQ-Tools-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\Admin\Documents\GitHub\BOQ-TOOLS\dist\BOQ-Tools.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\BOQ Tools"; Filename: "{app}\BOQ-Tools.exe"
Name: "{commondesktop}\BOQ Tools"; Filename: "{app}\BOQ-Tools.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\BOQ-Tools.exe"; Description: "{cm:LaunchProgram,BOQ Tools}"; Flags: nowait postinstall skipifsilent
