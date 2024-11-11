[Setup]
AppName=Folder Creation Tool
AppVersion=1.0
DefaultDirName={autopf}\CreateFolders
DefaultGroupName=CreateFolders
OutputDir=.\Output
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Files]
; Add your EXE file to the setup package
Source: "C:\Users\stefa\Desktop\APP DESKTOP CREATE FOLDER\create_folders.exe"; DestDir: "{app}"; Flags: ignoreversion
; Optionally, add the icon as a separate file (not necessary if the icon is embedded in the EXE)
Source: "C:\Users\stefa\Desktop\APP DESKTOP CREATE FOLDER\create_folders.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Create a desktop shortcut for the application with the correct icon
Name: "{autoprograms}\Folder Creation Tool"; Filename: "{app}\create_folders.exe"; WorkingDir: "{app}"; IconFilename: "{app}\create_folders.ico"

[Run]
; Automatically run the EXE file after installation
Filename: "{app}\create_folders.exe"; StatusMsg: "Running Create Folders App..."; Flags: nowait postinstall skipifsilent
