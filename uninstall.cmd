@echo off
setlocal
rem 脚本是 UTF-8：切到 65001，否则中文提示在中文系统（936）下是乱码
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  PPT 触控单击翻页 - 卸载
echo ============================================================
echo.

rem 组件注册在 HKCU\Software\Classes（按用户），卸载同样不需要管理员权限。
echo [1/3] 退出代理...
taskkill /im PptTouchClick.exe /f >nul 2>&1

echo [2/3] 取消开机启动...
if exist "PptTouchClick.exe" "PptTouchClick.exe" --remove-autostart

echo [3/3] 注销组件...
if exist "%WINDIR%\SysWOW64\regsvr32.exe" (
    if exist "PptSlideProc.x64.dll" "%WINDIR%\System32\regsvr32.exe" /s /u "%CD%\PptSlideProc.x64.dll"
) else (
    if exist "PptSlideProc.ocx" "%WINDIR%\System32\regsvr32.exe" /s /u "%CD%\PptSlideProc.ocx"
)
rem 兜底：直接删掉当前用户的 COM 注册
reg delete "HKCU\Software\Classes\CLSID\{2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21}" /f >nul 2>&1
reg delete "HKCU\Software\Classes\PptTouchClick.SlideWatcher" /f >nul 2>&1
reg delete "HKCU\Software\Classes\Interface\{7D2E1B44-03C5-4F88-B172-9E4D1860C29F}" /f >nul 2>&1

echo.
echo 已完成。可删除本文件夹。
echo 设置项（HKCU\Software\PptTouchClick）如不再需要，可在 regedit 里删除。
echo.
pause
