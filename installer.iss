; Inno Setup script: wraps the PyInstaller output into PrepInvoices-Setup.exe
[Setup]
AppName=Prep Invoices
AppVersion=1.0
DefaultDirName={autopf}\Prep Invoices
DefaultGroupName=Prep Invoices
OutputDir=installer
OutputBaseFilename=PrepInvoices-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\PrepInvoices.exe

[Files]
Source: "dist\PrepInvoices\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{autoprograms}\Prep Invoices"; Filename: "{app}\PrepInvoices.exe"
Name: "{autodesktop}\Prep Invoices"; Filename: "{app}\PrepInvoices.exe"

[Run]
Filename: "{app}\PrepInvoices.exe"; Description: "Open Prep Invoices now"; Flags: nowait postinstall skipifsilent
