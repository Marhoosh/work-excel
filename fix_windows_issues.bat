@echo off
echo ======================================
echo Excel数据处理工具 - Windows问题修复工具
echo ======================================
echo.
echo 本工具将尝试修复以下常见问题:
echo  1. 权限问题
echo  2. DLL丢失问题
echo  3. 程序运行时闪退问题
echo  4. 防病毒软件拦截问题
echo.

:menu
echo 请选择要执行的操作:
echo  [1] 检查和修复运行环境
echo  [2] 重新安装必要的VC++运行库
echo  [3] 为应用程序添加防病毒软件排除项
echo  [4] 修复打包后的程序
echo  [5] 清理临时文件和缓存
echo  [0] 退出
echo.

set /p choice=请输入选项(0-5): 

if "%choice%"=="1" goto check_env
if "%choice%"=="2" goto install_vcpp
if "%choice%"=="3" goto add_exclusion
if "%choice%"=="4" goto fix_app
if "%choice%"=="5" goto clean_temp
if "%choice%"=="0" goto end

echo 无效的选项，请重新选择。
goto menu

:check_env
echo.
echo 正在检查运行环境...
echo.

REM 检查是否以管理员权限运行
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [警告] 当前脚本未以管理员权限运行
    echo         某些修复操作可能无法完成
    echo         建议右键点击此脚本，选择"以管理员身份运行"
    echo.
    pause
)

REM 检查系统架构
echo 系统架构: 
if defined PROCESSOR_ARCHITEW6432 (
    echo 64位系统 (运行在32位模式下)
) else if defined PROCESSOR_ARCHITECTURE (
    if %PROCESSOR_ARCHITECTURE%==AMD64 (
        echo 64位系统
    ) else (
        echo 32位系统
    )
) else (
    echo 未知
)

REM 检查Windows版本
ver
echo.

REM 检查.NET Framework
echo 检查 .NET Framework...
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Version >nul 2>&1
if %errorlevel% equ 0 (
    echo [成功] 已安装 .NET Framework 4.x
) else (
    echo [警告] 未检测到 .NET Framework 4.x，某些功能可能无法使用
)
echo.

echo 环境检查完成。按任意键返回主菜单...
pause >nul
goto menu

:install_vcpp
echo.
echo 正在下载并安装 Visual C++ 运行库...
echo 这可能需要几分钟时间，请耐心等待
echo.

set VC2013_URL=https://aka.ms/highdpimfc2013x64
set VC2015_2019_URL=https://aka.ms/vs/16/release/vc_redist.x64.exe

REM 下载VC++ 2015-2019运行库
echo 下载 Visual C++ 2015-2019 运行库...
bitsadmin /transfer vcredist /download /priority normal "%VC2015_2019_URL%" "%CD%\vc_redist.exe"

if exist "%CD%\vc_redist.exe" (
    echo 正在安装 Visual C++ 2015-2019 运行库...
    start /wait vc_redist.exe /passive /norestart
    if %errorlevel% equ 0 (
        echo [成功] Visual C++ 2015-2019 运行库已安装
    ) else (
        echo [错误] 安装失败，错误代码: %errorlevel%
    )
    del /f /q vc_redist.exe
) else (
    echo [错误] 下载失败，请检查网络连接或手动下载:
    echo %VC2015_2019_URL%
)

echo.
echo 操作完成。按任意键返回主菜单...
pause >nul
goto menu

:add_exclusion
echo.
echo 添加杀毒软件排除项...
echo.

REM 确认是否以管理员身份运行
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 此操作需要管理员权限
    echo         请右键点击此脚本，选择"以管理员身份运行"
    echo.
    pause
    goto menu
)

REM 获取应用程序路径
set /p app_path=请输入Excel数据处理工具.exe的完整路径: 

if not exist "%app_path%" (
    echo [错误] 找不到指定的文件
    pause
    goto menu
)

REM 添加到Windows Defender排除项
echo 正在添加到Windows Defender排除项...
powershell -Command "Add-MpPreference -ExclusionPath '%app_path%'" >nul 2>&1
if %errorlevel% equ 0 (
    echo [成功] 已添加排除项到Windows Defender
) else (
    echo [警告] 无法添加排除项到Windows Defender，可能不是默认防病毒软件
)

echo.
echo 如果您使用其他杀毒软件，请手动将以下文件添加到排除列表:
echo %app_path%
echo.
echo 操作完成。按任意键返回主菜单...
pause >nul
goto menu

:fix_app
echo.
echo 修复打包后的程序...
echo.

REM 获取应用程序路径
set /p app_path=请输入Excel数据处理工具.exe的完整路径: 

if not exist "%app_path%" (
    echo [错误] 找不到指定的文件
    pause
    goto menu
)

REM 检查应用程序文件
echo 检查应用程序文件完整性...
set app_dir=%~dp0
if exist "%app_dir%\dist\Excel数据处理工具.exe" (
    echo [信息] 检测到打包文件，正在修复...
    
    REM 复制新的可执行文件
    copy /y "%app_dir%\dist\Excel数据处理工具.exe" "%app_path%" >nul 2>&1
    if %errorlevel% equ 0 (
        echo [成功] 已替换可执行文件
    ) else (
        echo [错误] 无法复制文件，可能是权限问题或文件正在使用
    )
) else (
    echo [警告] 未检测到源打包文件，无法进行修复
    echo         请重新运行build_windows.bat生成打包文件
)

echo.
echo 操作完成。按任意键返回主菜单...
pause >nul
goto menu

:clean_temp
echo.
echo 清理临时文件和缓存...
echo.

REM 清理可能的临时文件
if exist "build" rmdir /s /q "build"
if exist "__pycache__" rmdir /s /q "__pycache__"
if exist "*.pyc" del /f /q "*.pyc"
if exist "*.pyo" del /f /q "*.pyo"
if exist "*.spec" attrib -r -h -s "*.spec"

REM 清理用户临时文件中可能的缓存
echo 清理用户临时文件夹中的相关缓存...
del /f /q "%TEMP%\Excel数据处理工具*" >nul 2>&1
del /f /q "%TEMP%\pip-*" >nul 2>&1

echo [成功] 临时文件清理完成

echo.
echo 操作完成。按任意键返回主菜单...
pause >nul
goto menu

:end
echo.
echo 感谢使用Excel数据处理工具修复工具!
echo.
echo 如果问题仍未解决，请联系开发者获取支持。
pause 