@echo off
chcp 65001 >nul
title 美的对账助手 - 桌面快捷方式安装器

echo ========================================================
echo   📦 美的对账助手 - 自动化桌面安装工具
echo ========================================================
echo.

:: 获取当前脚本所在绝对路径（去除末尾反斜杠）
set "CURRENT_DIR=%~dp0"
if "%CURRENT_DIR:~-1%"=="\" set "CURRENT_DIR=%CURRENT_DIR:~0,-1%"

set "SCRIPT_FILE=%CURRENT_DIR%\meidi_console.pyw"
set "ICON_FILE=%CURRENT_DIR%\app_icon.ico"
set "ENV_FILE=%CURRENT_DIR%\.env"
set "ENV_EXAMPLE=%CURRENT_DIR%\.env.example"

:: 检查 .env 是否存在，若无则自动根据模板创建
if not exist "%ENV_FILE%" (
    if exist "%ENV_EXAMPLE%" (
        echo [提示] 检测到首次运行，已自动从 .env.example 生成 .env 模版文件。
        copy "%ENV_EXAMPLE%" "%ENV_FILE%" >nul
    )
)

echo [1/3] 正在通过 Windows 原生 COM 接口生成桌面专属快捷方式...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ws = New-Object -ComObject WScript.Shell; " ^
    "$curDir = '%CURRENT_DIR%'; " ^
    "$scriptPath = '%SCRIPT_FILE%'; " ^
    "$iconPath = '%ICON_FILE%'; " ^
    "$desktopPath = [Environment]::GetFolderPath('Desktop'); " ^
    "$lnkPath = Join-Path $desktopPath '美的对账助手.lnk'; " ^
    "$s = $ws.CreateShortcut($lnkPath); " ^
    "$s.TargetPath = 'pythonw.exe'; " ^
    "$s.Arguments = ('\"' + $scriptPath + '\"'); " ^
    "$s.WorkingDirectory = $curDir; " ^
    "$s.IconLocation = ($iconPath + ',0'); " ^
    "$s.Hotkey = 'Ctrl+Alt+M'; " ^
    "$s.Description = '美的仓储自动化控制台'; " ^
    "$s.Save(); " ^
    "$startMenuLnk = Join-Path ([Environment]::GetFolderPath('Programs')) '美的对账助手.lnk'; " ^
    "Copy-Item -Path $lnkPath -Destination $startMenuLnk -Force; " ^
    "Write-Host '   ✅ 快捷方式已成功部署至桌面与开始菜单！'"

echo [2/3] 刷新 Windows 桌面图标缓存...
ie4uinit.exe -show 2>nul

echo [3/3] 正在启动控制台...
start "" pythonw.exe "%SCRIPT_FILE%"

echo.
echo ========================================================
echo   🎉 安装完成！
echo   - 桌面已生成带有现代流光 M 图标的【美的对账助手】
echo   - 全局快捷键已生效：Ctrl + Alt + M
echo ========================================================
echo.
timeout /t 3 /nobreak >nul
exit
