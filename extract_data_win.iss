[Setup]
AppName=ExtractData
AppVersion=1.2
DefaultDirName={pf}\ExtractData
DefaultGroupName=ExtractData
OutputBaseFilename=ExtractData_Setup
Compression=lzma
SolidCompression=yes

[Files]
; 将C#生成的可执行文件和DLL文件添加到安装包
Source: "C:\Users\xieli\source\repos\extract_data_win\bin\Release\extract_data_win.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\xieli\source\repos\extract_data_win\bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
Source: "C:\Users\xieli\source\repos\extract_data_win\bin\Release\*.config"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

; 将Python脚本生成的可执行文件添加到安装包
Source: "C:\Users\xieli\source\repos\extract_data_win\py\dist\extract_data.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 创建开始菜单快捷方式
Name: "{group}\ExtractData"; Filename: "{app}\extract_data_win.exe"
; 创建桌面快捷方式
Name: "{commondesktop}\ExtractData"; Filename: "{app}\extract_data_win.exe"

[Run]
; 安装后运行应用程序
Filename: "{app}\extract_data_win.exe"; Description: "{cm:LaunchProgram,MyApp}"; Flags: nowait postinstall skipifsilent
