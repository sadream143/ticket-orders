; Inno Setup script - builds TicketOrders-Setup-<version>.exe
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={{942ACD9A-F40F-4684-BF8C-ADFDC4447B50}
AppName=Ticket Orders
AppVersion={#AppVersion}
AppPublisher=Ticket Orders
DefaultDirName={localappdata}\Programs\TicketOrders
DefaultGroupName=Ticket Orders
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\release
OutputBaseFilename=TicketOrders-Setup-{#AppVersion}
SetupIconFile=..\build\app.ico
UninstallDisplayIcon={app}\TicketOrders.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
AppMutex=TicketOrdersAppMutex
CloseApplications=yes
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[InstallDelete]
; remove old program files before copying the new ones (your data in Documents is never touched)
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\tesseract"

[Files]
Source: "..\dist\TicketOrders\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\Ticket Orders"; Filename: "{app}\TicketOrders.exe"
Name: "{autodesktop}\Ticket Orders"; Filename: "{app}\TicketOrders.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\TicketOrders.exe"; Description: "Launch Ticket Orders"; Flags: nowait postinstall
