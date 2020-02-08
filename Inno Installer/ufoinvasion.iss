; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[ISSI]

#define ISSI_IncludePath "C:\Archivos de programa\ISSI"

#define ISSI_URL "http://www.davidgf.net"
#define ISSI_UrlText "www.davidgf.net"

#include "_issi.isi"

[Setup]
PrivilegesRequired=admin
AppID=UFOInvasion
AppName=UFO Invasion
AppVerName=UFO Invasion BETA
AppPublisher=David GF Games
AppPublisherURL=http://www.davidgf.net
AppSupportURL=http://www.davidgf.net
AppUpdatesURL=http://www.davidgf.net
DefaultDirName={pf}\David GF Games\UFO Invasion
DefaultGroupName=UFO Invasion
LicenseFile=licence.rtf
OutputDir=..\
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes
VersionInfoVersion=0.1.0.0

;Compression=none
;SolidCompression=no

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked


[Files]

Source: "..\ufoinvasion.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\data.dat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Game.dat"; DestDir: "{app}"; Flags: ignoreversion

; ----- VB 6 ------
Source: "vb6\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace sharedfile regserver uninsneveruninstall
Source: "vb6\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "vb6\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,5; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vb6\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,5; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vb6\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,5; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vb6\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,4; Flags: restartreplace uninsneveruninstall sharedfile regserver


; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\UFO Invasion"; Filename: "{app}\ufoinvasion.exe"
Name: "{group}\{cm:UninstallProgram,UFO Invasion}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\UFO Invasion"; Filename: "{app}\ufoinvasion.exe"; Tasks: desktopicon

