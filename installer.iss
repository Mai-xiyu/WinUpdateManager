[Setup]
; 唯一标识符，用于卸载和更新
AppId={{8A3B2194-1234-4B56-A789-0123456789AB}
AppName=Windows 更新管理工具
#ifndef AppVersion
#define AppVersion "1.0.0"
#endif
AppVersion={#AppVersion}
AppPublisher=WinUpdateManager
; 默认安装目录 (Program Files)
DefaultDirName={autopf}\WinUpdateManager
; 默认开始菜单文件夹名
DefaultGroupName=Windows 更新管理工具
; 禁用选择开始菜单文件夹页面（简化安装）
DisableProgramGroupPage=yes
; 允许没有卸载图标
AllowNoIcons=yes
; 输出目录
OutputDir=.\Output
; 输出的安装包文件名
OutputBaseFilename=WinUpdateManager_Setup
; 安装程序的图标
SetupIconFile=WinUpdateManager\app.ico
; 卸载程序的图标
UninstallDisplayIcon={app}\WinUpdateManager.exe
; 压缩算法
Compression=lzma2
SolidCompression=yes
; 现代向导样式
WizardStyle=modern
; 64位安装模式
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Default.isl,ChineseSimplified.isl"

[Tasks]
; 默认勾选创建桌面快捷方式
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
; 将 publish 目录下的所有文件打包进去
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 开始菜单快捷方式
Name: "{autoprograms}\Windows 更新管理工具"; Filename: "{app}\WinUpdateManager.exe"
; 桌面快捷方式
Name: "{autodesktop}\Windows 更新管理工具"; Filename: "{app}\WinUpdateManager.exe"; Tasks: desktopicon

[Run]
; 安装完成后提供运行程序的选项
Filename: "{app}\WinUpdateManager.exe"; Description: "{cm:LaunchProgram,Windows 更新管理工具}"; Flags: nowait postinstall skipifsilent shellexec
