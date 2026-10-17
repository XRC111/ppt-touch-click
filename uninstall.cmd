@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  PPT 触控单击翻页 - 卸载
echo ============================================================
echo.

net session >nul 2>&1
if errorlevel 1 (
    echo [错误] 需要管理员权限。请右键本文件，选择“以管理员身份运行”。
    pause
    exit /b 1
)

echo [1/3] 退出代理...
taskkill /im PptTouchClick.exe /f >nul 2>&1

echo [2/3] 取消开机启动...
if exist "PptTouchClick.exe" "PptTouchClick.exe" --remove-autostart

echo [3/3] 注销组件...
if exist "PptSlideProc.x64.dll" regsvr32 /s /u "%CD%\PptSlideProc.x64.dll"
if exist "PptSlideProc.ocx"     regsvr32 /s /u "%CD%\PptSlideProc.ocx"

echo.
echo 已完成。可删除本文件夹。
echo 设置项（HKCU\Software\PptTouchClick）如不再需要，可在 regedit 里删除。
echo.
pause
