; 万能办公助手 v6.2 — Inno Setup 安装脚本
; 用 Inno Setup 6 (https://jrsoftware.org/isinfo.php) 编译
; ISCC.exe "万能办公助手v6.2商业版安装向导.iss"

#define MyAppName "万能办公助手"
#define MyAppVersion "6.2"
#define MyAppPublisher "万能办公助手团队"
#define MyAppURL "https://github.com/anganglijinliang/office-assistant"
#define MyAppExeName "万能办公助手v6.2商业版.exe"

[Setup]
AppId={{B8A3C1E2-4F5D-4A6E-8B7C-9D0E1F2A3B4C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=
InfoBeforeFile=
OutputDir=.\dist
OutputBaseFilename=万能办公助手v6.2安装程序
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName} v{#MyAppVersion}
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式:"; Flags: checkedonce

[Files]
Source: "dist\\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; 单文件模式（onefile），无需打包 _internal 目录
; 如需非单文件构建，取消下面行的注释并注释掉上面那行
; Source: "dist\\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Source: "dist\\_internal\\*"; DestDir: "{app}\\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: postinstall nowait skipifsilent

[UninstallRun]
Filename: "{app}\{#MyAppExeName}"; Parameters: "--uninstall"; Flags: runhidden waituntilterminated

[Code]
function InitializeUninstall: Boolean;
begin
  Result := MsgBox('确定要卸载「万能办公助手 v6.2」吗？' + #13#10 + #13#10 +
    '此操作将删除所有程序文件和配置数据。', mbConfirmation, MB_YESNO) = idYes;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // 清理 APPDATA 配置
    DelTree(ExpandConstant('{userappdata}\OfficeAssistant_v6.2'), True, True, True);
  end;
end;
