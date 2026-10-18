@echo off
setlocal
rem 脚本是 UTF-8：切到 65001，否则中文提示在中文系统（936）下是乱码
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  PPT 触控单击翻页 - 安装
echo ============================================================
echo.

net session >nul 2>&1
if errorlevel 1 (
    echo [错误] 需要管理员权限。
    echo        请右键本文件，选择“以管理员身份运行”。
    echo.
    pause
    exit /b 1
)

if not exist "PptSlideProc.x64.dll" (
    echo [错误] 同目录下找不到 PptSlideProc.x64.dll，请把压缩包完整解压后再运行。
    pause
    exit /b 1
)
if not exist "PptTouchClick.exe" (
    echo [错误] 同目录下找不到 PptTouchClick.exe。
    pause
    exit /b 1
)

echo [1/5] 注册页码/触摸组件...
rem 64 位系统只注册 x64 组件（主程序是 64 位）；32 位系统注册 ocx。
rem 注意：HKCU\Software\Classes\CLSID 不区分 32/64 位，两个都注册会互相覆盖，
rem 而且 64 位 regsvr32 也加载不了 32 位 DLL。
if exist "%WINDIR%\SysWOW64\regsvr32.exe" (
    "%WINDIR%\System32\regsvr32.exe" /s "%CD%\PptSlideProc.x64.dll"
) else (
    "%WINDIR%\System32\regsvr32.exe" /s "%CD%\PptSlideProc.ocx"
)
if errorlevel 1 echo        （regsvr32 返回错误，详见同目录 PptSlideProc.log）

echo [2/5] 检查注册结果...
reg query "HKCU\Software\Classes\CLSID\{2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21}\InprocServer32" >nul 2>&1
if errorlevel 1 (
    echo        [警告] 组件未注册成功：滑动过滤不可用，但“点击空白翻页”仍然有效。
) else (
    echo        组件注册成功。
)

echo [3/5] 设置开机自动启动...
"PptTouchClick.exe" --install-autostart
if errorlevel 1 echo        （设置失败，可稍后在托盘菜单里手动设置）

echo [4/5] 启动代理...
taskkill /im PptTouchClick.exe /f >nul 2>&1
start "" "PptTouchClick.exe"

echo [5/5] 完成。
echo.
echo   托盘里会出现黄色播放图标；右键可查看状态 / 暂停 / 打开日志。
echo   放映时：点空白处 = 下一页；点文字、图片、按钮 = 不翻页。
echo.
echo   卸载：右键 uninstall.cmd -> 以管理员身份运行。
echo.
pause
