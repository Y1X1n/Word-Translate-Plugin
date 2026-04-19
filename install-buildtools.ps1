# 自动下载并安装 Build Tools for Visual Studio 2022

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  安装 Build Tools for Visual Studio 2022" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 下载路径
$downloadUrl = "https://aka.ms/vs/17/release/vs_buildtools.exe"
$installerPath = "$env:TEMP\vs_buildtools.exe"

# 检查是否已安装
$msbuildPaths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)

$alreadyInstalled = $false
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        Write-Host "✓ MSBuild 已安装: $path" -ForegroundColor Green
        $alreadyInstalled = $true
        break
    }
}

if ($alreadyInstalled) {
    Write-Host ""
    Write-Host "Build Tools 已安装，无需重复安装。" -ForegroundColor Green
    Write-Host ""
    Read-Host "按 Enter 键退出"
    exit 0
}

# 下载安装程序
Write-Host "正在下载 Build Tools 安装程序..." -ForegroundColor Yellow
Write-Host "下载地址: $downloadUrl" -ForegroundColor Gray

try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing
    Write-Host "✓ 下载完成: $installerPath" -ForegroundColor Green
} catch {
    Write-Host "✗ 下载失败: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "请手动下载并安装:" -ForegroundColor Yellow
    Write-Host "https://visualstudio.microsoft.com/downloads/" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "按 Enter 键退出"
    exit 1
}

# 运行安装程序
Write-Host ""
Write-Host "正在安装 Build Tools..." -ForegroundColor Yellow
Write-Host "这需要几分钟时间，请耐心等待..." -ForegroundColor Gray
Write-Host ""

# 安装参数
$arguments = @(
    "--quiet"                                          # 静默安装
    "--wait"                                           # 等待安装完成
    "--add", "Microsoft.VisualStudio.Workload.MSBuildTools"  # MSBuild
    "--add", "Microsoft.VisualStudio.Workload.NetCoreBuildTools"  # .NET Core
    "--add", "Microsoft.VisualStudio.Component.NuGet"  # NuGet
    "--add", "Microsoft.VisualStudio.Component.Roslyn.Compiler"  # C# 编译器
    "--add", "Microsoft.Net.Component.4.8.TargetingPack"  # .NET 4.8
    "--add", "Microsoft.Net.Component.4.8.SDK"          # .NET 4.8 SDK
)

Write-Host "安装命令: $installerPath $($arguments -join ' ')" -ForegroundColor Gray
Write-Host ""

try {
    $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  安装成功!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        
        # 添加到 PATH
        $msbuildPath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin"
        if (Test-Path $msbuildPath) {
            $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
            if (-not $currentPath.Contains($msbuildPath)) {
                [Environment]::SetEnvironmentVariable("Path", "$currentPath;$msbuildPath", "User")
                Write-Host "✓ 已添加到用户 PATH: $msbuildPath" -ForegroundColor Green
            }
        }
        
        Write-Host ""
        Write-Host "请重新打开 PowerShell 或命令提示符，然后运行:" -ForegroundColor Yellow
        Write-Host "  .\build-standalone.bat" -ForegroundColor Cyan
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "  安装失败 (退出代码: $($process.ExitCode))" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host ""
        Write-Host "请尝试手动安装:" -ForegroundColor Yellow
        Write-Host "1. 访问 https://visualstudio.microsoft.com/downloads/" -ForegroundColor Cyan
        Write-Host "2. 下载 'Build Tools for Visual Studio 2022'" -ForegroundColor Cyan
        Write-Host "3. 安装 '.NET 桌面生成工具' 工作负载" -ForegroundColor Cyan
        Write-Host ""
    }
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  安装出错" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "错误信息: $_" -ForegroundColor Red
    Write-Host ""
}

# 清理
if (Test-Path $installerPath) {
    Remove-Item $installerPath -Force
}

Read-Host "按 Enter 键退出"
