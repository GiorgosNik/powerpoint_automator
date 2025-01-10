[Setup]
AppName=GEP Weather to Video      
OutputDir=../../dist
AppVersion=1.0.0
OutputBaseFilename=GEP Weather to Video Installer
DefaultDirName={autopf}\GEP Weather to Video    
Compression=lzma
SolidCompression=yes

[Files]
Source: "../../dist/GEP Weather to Video.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "../../assets/weather-news.ico"; DestDir: "{app}/assets"; Flags: ignoreversion
Source: "../template.pptx"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional tasks"; Flags: unchecked

[Icons]
Name: "{group}\GEP Weather to Video"; Filename: "{app}\GEP Weather to Video.exe"
Name: "{autodesktop}\GEP Weather to Video"; Filename: "{app}\GEP Weather to Video.exe"; IconFilename: "{app}/assets/weather-news.ico"; Tasks: desktopicon