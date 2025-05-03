@echo off
echo ===================================
echo 开始构建Excel数据处理工具...
echo ===================================

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] Python未安装或不在PATH环境变量中
    pause
    exit /b 1
)

REM 先运行环境检查脚本
echo [信息] 检查环境...
python check_environment.py
if %errorlevel% neq 0 (
    echo [警告] 环境检查未通过，但将继续尝试打包
    echo         如果打包失败，请检查并解决上述问题
    echo.
    pause
)

REM 创建虚拟环境（可选）
echo [信息] 创建虚拟环境...
if exist "venv" (
    echo [信息] 使用已存在的虚拟环境
) else (
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [错误] 创建虚拟环境失败，尝试不使用虚拟环境继续...
        set USE_VENV=0
    ) else (
        set USE_VENV=1
    )
)

REM 尝试激活虚拟环境
if "%USE_VENV%"=="1" (
    echo [信息] 激活虚拟环境...
    call venv\Scripts\activate
)

REM 安装依赖
echo [信息] 安装依赖...
python -m pip install --upgrade pip
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 安装依赖失败
    pause
    exit /b 1
)

REM 检查dist目录并清理
if exist "dist" (
    echo [信息] 清理旧的打包文件...
    rmdir /s /q "dist"
)
if exist "build" (
    echo [信息] 清理构建目录...
    rmdir /s /q "build"
)

REM 开始打包
echo [信息] 开始打包Windows可执行程序...
pyinstaller --noconfirm excel-app.spec
if %errorlevel% neq 0 (
    echo [错误] 打包过程中出现错误
    pause
    exit /b 1
)

REM 检查生成的文件
if not exist "dist\Excel数据处理工具.exe" (
    echo [错误] 打包完成，但未找到生成的可执行文件
    echo         请检查是否有权限问题或磁盘空间不足
    pause
    exit /b 1
)

echo [成功] 打包完成，输出文件位于 dist 目录
echo.
echo       文件：dist\Excel数据处理工具.exe
echo       大小：%~z0 字节
echo.
echo 按任意键后将打开dist目录...
pause

REM 打开文件资源管理器显示dist目录
start explorer dist

REM 如果使用了虚拟环境，尝试退出
if "%USE_VENV%"=="1" (
    call venv\Scripts\deactivate
) 