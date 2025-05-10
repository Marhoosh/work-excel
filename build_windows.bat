@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ============================================================
echo               Excel数据处理工具 - Windows打包脚本
echo ============================================================
echo.

REM 检查Python安装
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python安装!
    echo 请安装Python 3.6或更高版本，并确保添加到PATH中
    echo 可以从 https://www.python.org/downloads/ 下载
    goto :error
)

REM 创建虚拟环境
echo 正在创建虚拟环境...
if exist build_env rmdir /s /q build_env
python -m venv build_env
if %errorlevel% neq 0 (
    echo 警告: 创建虚拟环境失败，将使用全局Python环境
) else (
    echo 激活虚拟环境...
    call build_env\Scripts\activate
)

REM 安装/升级pip
echo 升级pip...
python -m pip install --upgrade pip

REM 安装必要的依赖
echo 安装依赖...
python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/
if %errorlevel% neq 0 (
    echo 警告: 安装requirements.txt中的依赖失败
    echo 尝试单独安装每个依赖...
    python -m pip install openpyxl
    python -m pip install pandas
    python -m pip install numpy
    python -m pip install xlrd
    REM tkinter通常作为Python标准库的一部分，通常不需要单独安装
)

REM 检查tkinter是否可用
echo 检查tkinter是否可用...
python -c "import tkinter; print('tkinter检查: 已安装')" 2>nul
if %errorlevel% neq 0 (
    echo 警告: tkinter未安装或无法导入!
    echo tkinter是Python标准库的一部分，无法通过pip安装
    echo 请确保安装了完整版的Python，并在安装时勾选了"tcl/tk和IDLE"选项
    echo 继续打包，但如果程序依赖tkinter，最终的可执行文件可能无法运行
    echo.
)

REM 安装PyInstaller
echo 安装PyInstaller...
python -m pip install pyinstaller
if %errorlevel% neq 0 (
    echo 错误: 安装PyInstaller失败!
    goto :error
)

REM 清除旧的构建文件
echo 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__

REM 确保有spec文件
if not exist excel-app.spec (
    echo 错误: 缺少excel-app.spec文件!
    goto :error
)

REM 运行PyInstaller
echo 开始打包应用程序...
pyinstaller --noconfirm excel-app.spec
if %errorlevel% neq 0 (
    echo 错误: PyInstaller打包失败!
    goto :error
)

REM 检查输出文件是否存在
if not exist dist\Excel数据处理工具.exe (
    echo 错误: 未找到构建的可执行文件!
    echo 检查spec文件是否正确配置。
    goto :error
)

echo ============================================================
echo                       打包成功!
echo        可执行文件位于: dist\Excel数据处理工具.exe
echo ============================================================

REM 打开输出目录
start explorer.exe dist
goto :end

:error
echo ============================================================
echo                       打包失败!
echo ============================================================
echo 请检查错误消息，解决问题后重试。
pause
exit /b 1

:end
if exist build_env (
    echo 清理虚拟环境...
    call build_env\Scripts\deactivate
    REM 如果需要保留虚拟环境，可以注释掉下一行
    REM rmdir /s /q build_env
)
echo 按任意键退出...
pause > nul
exit /b 0 