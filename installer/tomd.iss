; ToMD 설치 파일 스크립트 (Inno Setup 6)
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppId={{8E6B1C4A-5D2F-4B7E-9A31-7C0D2E5F4A16}
AppName=ToMD 마크다운 변환기
AppVersion={#AppVersion}
AppPublisher=ToMD
DefaultDirName={localappdata}\Programs\ToMD
PrivilegesRequired=lowest
DisableDirPage=yes
DisableProgramGroupPage=yes
ChangesAssociations=yes
OutputDir=Output
OutputBaseFilename=ToMD-Setup
Compression=lzma2
SolidCompression=yes
UninstallDisplayName=ToMD 마크다운 변환기
UninstallDisplayIcon={app}\tomd.exe
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\dist\tomd\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[INI]
Filename: "{app}\version.ini"; Section: "tomd"; Key: "version"; String: "{#AppVersion}"

[UninstallDelete]
Type: files; Name: "{app}\version.ini"

[Registry]
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pdf\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwp\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwp\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwp\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwp\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwpx\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwpx\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwpx\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.hwpx\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsx\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xlsm\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.xls\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.csv\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.csv\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.csv\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.csv\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.tsv\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.tsv\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.tsv\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.tsv\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
Root: HKCU; Subkey: "Software\Classes\Directory\shell\ToMD"; ValueType: string; ValueName: ""; ValueData: "폴더 안 문서를 마크다운(MD)으로 변환"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Directory\shell\ToMD"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\tomd.exe"
Root: HKCU; Subkey: "Software\Classes\Directory\shell\ToMD"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"
Root: HKCU; Subkey: "Software\Classes\Directory\shell\ToMD\command"; ValueType: string; ValueName: ""; ValueData: """{app}\tomd.exe"" ""%1"""
