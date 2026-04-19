@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ==========================================
echo Word Translation Add-in 独立构建脚本
echo (无需 Visual Studio)
echo ==========================================
echo.

REM 设置变量
set "SOLUTION_FILE=WordTranslationAddin.sln"
set "PROJECT_FILE=WordTranslationAddin\WordTranslationAddin.csproj"
set "CONFIGURATION=Release"
set "PLATFORM=Any CPU"
set "OUTPUT_DIR=WordTranslationAddin\bin\Release"

REM 颜色代码
set "GREEN=[92m"
set "YELLOW=[93m"
set "RED=[91m"
set "RESET=[0m"

echo [步骤 1/6] 检查系统环境...
echo ==========================================

REM 检查操作系统版本
for /f "tokens=4-5 delims=. " %%i in ('ver') do set VERSION=%%i.%%j
if "%version%" == "10.0" (
    echo %GREEN%✓%RESET% Windows 10/11  detected
) else if "%version%" == "6.3" (
    echo %YELLOW%!%RESET% Windows 8.1 detected (可能需要额外配置)
) else if "%version%" == "6.1" (
    echo %YELLOW%!%RESET% Windows 7 detected (可能需要额外配置)
) else (
    echo %YELLOW%!%RESET% Unknown Windows version: %version%
)

REM 检查 .NET Framework
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if %errorlevel% == 0 (
    for /f "tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release ^| findstr Release') do (
        set "dotnetRelease=%%b"
        if !dotnetRelease! GEQ 528040 (
            echo %GREEN%✓%RESET% .NET Framework 4.8 or later installed (Release: !dotnetRelease!)
        ) else (
            echo %RED%✗%RESET% .NET Framework 4.8 not found. Please install from:
            echo   https://dotnet.microsoft.com/download/dotnet-framework/net48
            pause
            exit /b 1
        )
    )
) else (
    echo %RED%✗%RESET% .NET Framework 4.x not found
    pause
    exit /b 1
)

echo.
echo [步骤 2/6] 查找 MSBuild...
echo ==========================================

set "MSBUILD_PATH="

REM 尝试多个可能的 MSBuild 位置
set "PATHS_TO_CHECK="
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe;"
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe;"
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe;"
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe;"
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe;"
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe;"
set "PATHS_TO_CHECK=!PATHS_TO_CHECK!C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe;"

for %%p in (!PATHS_TO_CHECK!) do (
    if exist "%%p" (
        set "MSBUILD_PATH=%%p"
        echo %GREEN%✓%RESET% Found MSBuild at: %%p
        goto :msbuild_found
    )
)

REM 检查 PATH 中的 MSBuild
where msbuild >nul 2>&1
if %errorlevel% == 0 (
    for /f "tokens=*" %%a in ('where msbuild') do (
        set "MSBUILD_PATH=%%a"
        echo %GREEN%✓%RESET% Found MSBuild in PATH: %%a
        goto :msbuild_found
    )
)

echo %RED%✗%RESET% MSBuild not found!
echo.
echo Please install Build Tools for Visual Studio 2022:
echo   1. Download from: https://visualstudio.microsoft.com/downloads/
echo   2. Scroll to "Tools for Visual Studio" 
gecho   3. Select "Build Tools for Visual Studio 2022"
echo   4. Install workload: ".NET desktop build tools"
echo.
pause
exit /b 1

:msbuild_found
echo.

REM 显示 MSBuild 版本
echo MSBuild version:
"!MSBUILD_PATH!" -version | findstr "^[0-9]"
echo.

echo [步骤 3/6] 检查 NuGet...
echo ==========================================

set "NUGET_PATH="

REM 检查 PATH 中的 NuGet
where nuget >nul 2>&1
if %errorlevel% == 0 (
    for /f "tokens=*" %%a in ('where nuget') do (
        set "NUGET_PATH=%%a"
        echo %GREEN%✓%RESET% Found NuGet in PATH: %%a
        goto :nuget_found
    )
)

REM 检查常见位置
if exist "C:\Tools\nuget.exe" (
    set "NUGET_PATH=C:\Tools\nuget.exe"
    echo %GREEN%✓%RESET% Found NuGet at: C:\Tools\nuget.exe
    goto :nuget_found
)

if exist "%LOCALAPPDATA%\NuGet\nuget.exe" (
    set "NUGET_PATH=%LOCALAPPDATA%\NuGet\nuget.exe"
    echo %GREEN%✓%RESET% Found NuGet at: %LOCALAPPDATA%\NuGet\nuget.exe
    goto :nuget_found
)

echo %YELLOW%!%RESET% NuGet not found. Downloading...

REM 创建 Tools 目录
if not exist "C:\Tools" mkdir "C:\Tools"

REM 下载 NuGet
powershell -Command "Invoke-WebRequest -Uri 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile 'C:\Tools\nuget.exe'" >nul 2>&1

if exist "C:\Tools\nuget.exe" (
    set "NUGET_PATH=C:\Tools\nuget.exe"
    echo %GREEN%✓%RESET% NuGet downloaded to: C:\Tools\nuget.exe
    
    REM 添加到 PATH
    setx PATH "%PATH%;C:\Tools" >nul 2>&1
    echo %YELLOW%!%RESET% Added C:\Tools to PATH. Please restart command prompt to use nuget command globally.
) else (
    echo %RED%✗%RESET% Failed to download NuGet
    echo Please manually download from: https://www.nuget.org/downloads
    pause
    exit /b 1
)

:nuget_found
echo.

echo [步骤 4/6] 还原 NuGet 包...
echo ==========================================

echo Restoring packages for %SOLUTION_FILE%...
"!NUGET_PATH!" restore "%SOLUTION_FILE%" -Verbosity quiet

if %errorlevel% neq 0 (
    echo %RED%✗%RESET% Failed to restore NuGet packages
    echo Trying alternative method...
    
    REM 尝试使用 MSBuild 还原
    "!MSBUILD_PATH!" "%SOLUTION_FILE%" /t:Restore /p:Configuration=%CONFIGURATION% /verbosity:minimal
    
    if %errorlevel% neq 0 (
        echo %RED%✗%RESET% Package restore failed
        pause
        exit /b 1
    )
)

echo %GREEN%✓%RESET% Packages restored successfully
echo.

echo [步骤 5/6] 构建项目...
echo ==========================================

echo Building %SOLUTION_FILE%...
echo Configuration: %CONFIGURATION%
echo Platform: %PLATFORM%
echo.

"!MSBUILD_PATH!" "%SOLUTION_FILE%" ^
    /p:Configuration=%CONFIGURATION% ^
    /p:Platform="%PLATFORM%" ^
    /verbosity:minimal ^
    /consoleloggerparameters:Summary

if %errorlevel% neq 0 (
    echo.
    echo %RED%==========================================%RESET%
    echo %RED%  构建失败!%RESET%
    echo %RED%==========================================%RESET%
    echo.
    echo 常见错误及解决方案:
    echo   1. 缺少 Office Interop 程序集
    echo      - 安装 Microsoft Office 或 Office PIA
    echo   2. 缺少 .NET Framework 目标包
    echo      - 安装 .NET Framework 4.8 Developer Pack
    echo   3. 项目文件损坏
    echo      - 重新下载项目文件
    echo.
    pause
    exit /b 1
)

echo.
echo %GREEN%✓%RESET% Build completed successfully!
echo.

echo [步骤 6/6] 验证输出...
echo ==========================================

set "BUILD_SUCCESS=true"

if not exist "%OUTPUT_DIR%\WordTranslationAddin.dll" (
    echo %RED%✗%RESET% WordTranslationAddin.dll not found
    set "BUILD_SUCCESS=false"
) else (
    echo %GREEN%✓%RESET% WordTranslationAddin.dll
    for %%F in ("%OUTPUT_DIR%\WordTranslationAddin.dll") do (
        echo   Size: %%~zF bytes
        echo   Modified: %%~tF
    )
)

if not exist "%OUTPUT_DIR%\Newtonsoft.Json.dll" (
    echo %RED%✗%RESET% Newtonsoft.Json.dll not found
    set "BUILD_SUCCESS=false"
) else (
    echo %GREEN%✓%RESET% Newtonsoft.Json.dll
)

if not exist "%OUTPUT_DIR%\WordTranslationAddin.dll.config" (
    echo %YELLOW%!%RESET% WordTranslationAddin.dll.config not found (optional)
) else (
    echo %GREEN%✓%RESET% WordTranslationAddin.dll.config
)

echo.

if "%BUILD_SUCCESS%"=="true" (
    echo %GREEN%==========================================%RESET%
    echo %GREEN%  构建成功!%RESET%
    echo %GREEN%==========================================%RESET%
    echo.
    echo 输出目录: %CD%\%OUTPUT_DIR%
    echo.
    echo 下一步:
    echo   1. 运行 deploy.bat 安装插件
    echo   2. 或手动复制文件到 Word 插件目录
    echo   3. 重启 Microsoft Word
    echo.
    echo 安装命令:
    echo   deploy.bat
    echo.
) else (
    echo %RED%==========================================%RESET%
    echo %RED%  构建部分失败%RESET%
    echo %RED%==========================================%RESET%
    echo.
    echo 请检查错误信息并修复问题。
)

pause
