@echo off
chcp 65001 >nul
title 万能办公助手 v6.2 — 安装向导
echo ════════════════════════════════════════
echo   万能办公助手 v6.2 商业版 — 安装向导
echo ════════════════════════════════════════
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠ 请以「管理员身份运行」此安装程序！
    echo   右键本文件 → 以管理员身份运行
    echo.
    pause
    exit /b 1
)

:: 确定源 exe 路径
set "SCRIPT_DIR=%~dp0"
set "SOURCE_EXE=%SCRIPT_DIR%万能办公助手v6.2商业版.exe"

if not exist "%SOURCE_EXE%" (
    echo ❌ 未找到 万能办公助手v6.2商业版.exe
 echo    请将本文件与 exe 放在同一目录
    pause
    exit /b 1
)

:: 安装路径
set "INSTALL_DIR=%ProgramFiles%\万能办公助手"

echo 📍 安装路径: %INSTALL_DIR%
echo 📦 正在复制文件...

:: 创建安装目录
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: 复制 exe
copy /Y "%SOURCE_EXE%" "%INSTALL_DIR%\" >nul
if %errorlevel% neq 0 (
    echo ❌ 文件复制失败！请检查权限。
    pause
    exit /b 1
)

:: 创建开始菜单快捷方式
echo 📌 正在创建快捷方式...
set "STARTMENU=%AppData%\Microsoft\Windows\Start Menu\Programs\万能办公助手"
if not exist "%STARTMENU%" mkdir "%STARTMENU%"

:: 用 PowerShell 创建快捷方式
powershell -Command ^
    $WS = New-Object -ComObject WScript.Shell; ^
    $SC = $WS.CreateShortcut('%STARTMENU%\万能办公助手v6.2.lnk'); ^
    $SC.TargetPath = '%INSTALL_DIR%\万能办公助手v6.2商业版.exe'; ^
    $SC.WorkingDirectory = '%INSTALL_DIR%'; ^
    $SC.Description = '万能办公助手 v6.2 商业版'; ^
    $SC.Save(); ^
    $SC2 = $WS.CreateShortcut('%STARTMENU%\卸载万能办公助手.lnk'); ^
    $SC2.TargetPath = '%INSTALL_DIR%\卸载万能办公助手.bat'; ^
    $SC2.Save(); ^
    $SC3 = $WS.CreateShortcut('%USERPROFILE%\Desktop\万能办公助手v6.2.lnk'); ^
    $SC3.TargetPath = '%INSTALL_DIR%\万能办公助手v6.2商业版.exe'; ^
    $SC3.Description = '万能办公助手 v6.2 商业版'; ^
    $SC3.Save();

:: 复制卸载脚本到安装目录
copy /Y "%~f0" "%INSTALL_DIR%\卸载万能办公助手.bat" >nul

:: 写入卸载注册表（添加/删除程序）
echo 📝 正在注册卸载信息...
powershell -Command ^
    $path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\万能办公助手'; ^
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }; ^
    Set-ItemProperty -Path $path -Name 'DisplayName' -Value '万能办公助手 v6.2'; ^
    Set-ItemProperty -Path $path -Name 'DisplayVersion' -Value '6.2'; ^
    Set-ItemProperty -Path $path -Name 'Publisher' -Value '万能办公助手团队'; ^
    Set-ItemProperty -Path $path -Name 'InstallLocation' -Value '%INSTALL_DIR%'; ^
    Set-ItemProperty -Path $path -Name 'UninstallString' -Value '"%INSTALL_DIR%\卸载万能办公助手.bat"'; ^
    Set-ItemProperty -Path $path -Name 'DisplayIcon' -Value '%INSTALL_DIR%\万能办公助手v6.2商业版.exe'; ^
    Set-ItemProperty -Path $path -Name 'EstimatedSize' -Value 156000; ^
    Set-ItemProperty -Path $path -Name 'NoModify' -Value 1; ^
    Set-ItemProperty -Path $path -Name 'NoRepair' -Value 1;

echo.
echo ✅ 安装完成！
echo.
echo 📍 安装路径: %INSTALL_DIR%
echo 🏠 桌面快捷方式: 已创建
echo 📋 开始菜单: 已创建
echo ⚙ 控制面板「添加/删除程序」: 已注册
echo.
echo 是否现在启动程序？(Y/N)
set /p START_NOW=
if /i "%START_NOW%"=="Y" (
    start "" "%INSTALL_DIR%\万能办公助手v6.2商业版.exe"
)

pause
exit /b 0
