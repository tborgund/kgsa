; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{460E3C0E-6B13-4186-BA27-B4A9FEFDEFC0}
AppName=KGSA
AppVersion=1.17
;AppVerName=KGSA 1.17
AppPublisher=tborgund@gmail.com
AppPublisherURL=http://www.elkjop.no
AppSupportURL=http://www.elkjop.no
AppUpdatesURL=http://www.elkjop.no
DefaultDirName={pf}\KGSA
DefaultGroupName=KGSA
AllowNoIcons=yes
InfoBeforeFile=C:\Export\krav.txt
InfoAfterFile=C:\Export\KGSA les meg.pdf
OutputBaseFilename=setup
SetupIconFile=D:\Profil\Documents\Visual Studio 2012\Projects\KGSA\CVS_Parse\Ever.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Export\KGSA.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Export\FileHelpers.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Export\KGSA les meg.pdf"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\KGSA"; Filename: "{app}\KGSA.exe"
Name: "{group}\{cm:UninstallProgram,KGSA}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\KGSA"; Filename: "{app}\KGSA.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\KGSA.exe"; Description: "{cm:LaunchProgram,KGSA}"; Flags: nowait postinstall skipifsilent

