@echo off
chcp 65001 >nul
echo ==========================================
echo Word Translation Add-in 构建脚本
echo ==========================================
echo.

REM 设置变量
set "SOLUTION_FILE=WordTranslationAddin.sln"
set "CONFIGURATION=Release"
set "PLATFORM=Any CPU"

REM 检查 MSBuild
where msbuild >nul 2>nul
if %errorlevel% neq 0 (
    echo 错误: 找不到 MSBuild。请确保已安装 Visual Studio 并将其添加到系统 PATH。
    pause
    exit /b 1
)

echo 开始构建解决方案...
msbuild "%SOLUTION_FILE%" /p:Configuration=%CONFIGURATION% /p:Platform="%PLATFORM%" /verbosity:minimal

if %errorlevel% neq 0 (
    echo.
    echo 构建失败！
    pause
    exit /b 1
)

echo.
echo 构建成功！
echo.
echo 输出文件位置:
echo   - WordTranslationAddin.dll
echo   - WordTranslationAddin.dll.config
echo   - Newtonsoft.Json.dll
echo.

REM 检查是否需要构建安装程序
if "%1"=="--installer" (
    echo 正在构建安装程序...
    
    REM 检查 WiX
    where candle >nul 2>nul
    if %errorlevel% neq 0 (
        echo 警告: 找不到 WiX Toolset，跳过安装程序构建。
        echo 请安装 WiX Toolset 以构建安装程序。
    ) else (
        cd Installer
        msbuild Installer.wixproj /p:Configuration=%CONFIGURATION% /p:Platform=x86 /verbosity:minimal
        cd ..
        
        if %errorlevel% equ 0 (
            echo.
            echo 安装程序构建成功！
            echo 输出: Installer\bin\Release\WordTranslationAddin.msi
        ) else (
            echo.
            echo 安装程序构建失败！
        )
    )
)

echo.
echo 构建完成！
pause
