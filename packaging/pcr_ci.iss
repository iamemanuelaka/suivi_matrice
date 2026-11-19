; Installateur Windows de l'application PCR-CI (Inno Setup 6).
; Installation sans droits administrateur, dans le profil de l'utilisateur.
#define AppName "PCR-CI – Suivi des activités"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={{6F1C8E2A-4B7D-4C1E-9A55-3D2B7C1E0F41}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=PCR-CI
DefaultDirName={localappdata}\Programs\PCR-CI
DefaultGroupName=PCR-CI
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=PCR-CI-Installation-{#AppVersion}
SetupIconFile=icone.ico
UninstallDisplayIcon={app}\PCR-CI.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes

[Languages]
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "bureau"; Description: "Créer une icône sur le bureau"; GroupDescription: "Raccourcis :"

[Files]
Source: "..\dist\PCR-CI\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PCR-CI"; Filename: "{app}\PCR-CI.exe"
Name: "{group}\PCR-CI (dans le navigateur)"; Filename: "{app}\PCR-CI.exe"; Parameters: "--navigateur"
Name: "{userdesktop}\PCR-CI"; Filename: "{app}\PCR-CI.exe"; Tasks: bureau

[Run]
Filename: "{app}\PCR-CI.exe"; Description: "Lancer PCR-CI maintenant"; Flags: nowait postinstall skipifsilent
