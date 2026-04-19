# Word Translation Add-in 构建环境诊断脚本
# 用于检测构建环境并输出诊断报告

param(
    [switch]$Fix,
    [switch]$Verbose
)

$ErrorActionPreference = "Continue"
$script:Issues = @()
$script:Warnings = @()
$script:SuccessItems = @()

# 颜色输出函数
function Write-Success($message) {
    Write-Host "  [OK] $message" -ForegroundColor Green
    $script:SuccessItems += $message
}

function Write-Warning($message) {
    Write-Host "  [WARN] $message" -ForegroundColor Yellow
    $script:Warnings += $message
}

function Write-Error($message) {
    Write-Host "  [FAIL] $message" -ForegroundColor Red
    $script:Issues += $message
}

function Write-Info($message) {
    Write-Host "  [INFO] $message" -ForegroundColor Cyan
}

function Write-Section($title) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor White
    Write-Host "  $title" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor White
    Write-Host ""
}

# 检查 .NET Framework
function Test-DotNetFramework {
    Write-Section ".NET Framework 检查"
    
    $dotnetKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    
    if (Test-Path $dotnetKey) {
        $release = (Get-ItemProperty $dotnetKey -Name Release -ErrorAction SilentlyContinue).Release
        $version = (Get-ItemProperty $dotnetKey -Name Version -ErrorAction SilentlyContinue).Version
        
        if ($release -ge 528040) {
            Write-Success ".NET Framework 4.8 or later installed (Version: $version, Release: $release)"
        } elseif ($release -ge 461808) {
            Write-Warning ".NET Framework 4.7.2 installed (Version: $version). Recommend upgrading to 4.8"
        } else {
            Write-Error ".NET Framework version too old (Version: $version). Minimum 4.7.2 required"
            Write-Info "Download: https://dotnet.microsoft.com/download/dotnet-framework/net48"
        }
    } else {
        Write-Error ".NET Framework 4.x not found"
        Write-Info "Download: https://dotnet.microsoft.com/download/dotnet-framework/net48"
    }
}

# 检查 MSBuild
function Test-MSBuild {
    Write-Section "MSBuild 检查"
    
    $msbuildPaths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    )
    
    $found = $false
    foreach ($path in $msbuildPaths) {
        if (Test-Path $path) {
            $version = & $path -version 2>$null | Select-String "^\d+" | Select-Object -First 1
            Write-Success "MSBuild found at: $path"
            Write-Info "Version: $version"
            $found = $true
            break
        }
    }
    
    if (-not $found) {
        $msbuildInPath = Get-Command msbuild -ErrorAction SilentlyContinue
        if ($msbuildInPath) {
            $version = & msbuild -version 2>$null | Select-String "^\d+" | Select-Object -First 1
            Write-Success "MSBuild found in PATH: $($msbuildInPath.Source)"
            Write-Info "Version: $version"
            $found = $true
        }
    }
    
    if (-not $found) {
        Write-Error "MSBuild not found"
        Write-Info "Download Build Tools: https://visualstudio.microsoft.com/downloads/"
        Write-Info "Required workload: .NET desktop build tools"
    }
}

# 检查 NuGet
function Test-NuGet {
    Write-Section "NuGet 检查"
    
    $nuget = Get-Command nuget -ErrorAction SilentlyContinue
    
    if ($nuget) {
        $version = & nuget help 2>$null | Select-String "NuGet Version" | Select-Object -First 1
        Write-Success "NuGet found: $($nuget.Source)"
        Write-Info $version
    } else {
        $commonPaths = @(
            "C:\Tools\nuget.exe"
            "$env:LOCALAPPDATA\NuGet\nuget.exe"
            "$env:ProgramData\chocolatey\bin\nuget.exe"
        )
        
        $found = $false
        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                Write-Success "NuGet found at: $path"
                Write-Warning "NuGet not in PATH. Add '$([System.IO.Path]::GetDirectoryName($path))' to PATH"
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            Write-Error "NuGet not found"
            Write-Info "Download: https://www.nuget.org/downloads"
            
            if ($Fix) {
                Write-Info "Attempting to download NuGet..."
                try {
                    if (-not (Test-Path "C:\Tools")) {
                        New-Item -ItemType Directory -Path "C:\Tools" -Force | Out-Null
                    }
                    Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile "C:\Tools\nuget.exe"
                    Write-Success "NuGet downloaded to C:\Tools\nuget.exe"
                    
                    # Add to PATH
                    $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
                    if (-not $currentPath.Contains("C:\Tools")) {
                        [Environment]::SetEnvironmentVariable("Path", "$currentPath;C:\Tools", "User")
                        Write-Info "Added C:\Tools to User PATH"
                    }
                } catch {
                    Write-Error "Failed to download NuGet: $_"
                }
            }
        }
    }
}

# 检查 Office
function Test-Office {
    Write-Section "Microsoft Office 检查"
    
    $officePaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
    )
    
    $found = $false
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            $wordPath = (Get-ItemProperty $path -ErrorAction SilentlyContinue).'(default)'
            if ($wordPath -and (Test-Path $wordPath)) {
                $version = (Get-ItemProperty $path -ErrorAction SilentlyContinue).Version
                Write-Success "Microsoft Word found: $wordPath"
                Write-Info "Version: $version"
                $found = $true
                break
            }
        }
    }
    
    if (-not $found) {
        Write-Warning "Microsoft Office not found. Plugin requires Office to function."
        Write-Info "Note: Build can proceed without Office installed"
    }
}

# 检查 VSTO Runtime
function Test-VSTORuntime {
    Write-Section "VSTO Runtime 检查"
    
    $vstoKey = "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
    $vstoKey32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R"
    
    if (Test-Path $vstoKey) {
        $version = (Get-ItemProperty $vstoKey -Name Version -ErrorAction SilentlyContinue).Version
        Write-Success "VSTO Runtime installed (Version: $version)"
    } elseif (Test-Path $vstoKey32) {
        $version = (Get-ItemProperty $vstoKey32 -Name Version -ErrorAction SilentlyContinue).Version
        Write-Success "VSTO Runtime installed (x86, Version: $version)"
    } else {
        Write-Warning "VSTO Runtime not found"
        Write-Info "Download: https://www.microsoft.com/en-us/download/details.aspx?id=54251"
        Write-Info "Note: Required for plugin to run, but not for building"
    }
}

# 检查项目文件
function Test-ProjectFiles {
    Write-Section "项目文件检查"
    
    $requiredFiles = @(
        "WordTranslationAddin.sln"
        "WordTranslationAddin\WordTranslationAddin.csproj"
        "WordTranslationAddin\ThisAddIn.cs"
        "WordTranslationAddin\packages.config"
    )
    
    $allExist = $true
    foreach ($file in $requiredFiles) {
        if (Test-Path $file) {
            Write-Success "Found: $file"
        } else {
            Write-Error "Missing: $file"
            $allExist = $false
        }
    }
    
    if ($allExist) {
        Write-Info "All required project files present"
    }
}

# 检查磁盘空间
function Test-DiskSpace {
    Write-Section "磁盘空间检查"
    
    $drive = (Get-Location).Drive.Name
    $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$drive`:" 
    $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $totalSpaceGB = [math]::Round($disk.Size / 1GB, 2)
    
    Write-Info "Drive $drive`: $freeSpaceGB GB free of $totalSpaceGB GB"
    
    if ($freeSpaceGB -lt 1) {
        Write-Error "Low disk space: $freeSpaceGB GB remaining"
    } elseif ($freeSpaceGB -lt 5) {
        Write-Warning "Disk space getting low: $freeSpaceGB GB remaining"
    } else {
        Write-Success "Sufficient disk space available"
    }
}

# 生成诊断报告
function Show-DiagnosticReport {
    Write-Section "诊断报告"
    
    Write-Host "检查结果摘要:" -ForegroundColor White
    Write-Host "  通过: $($SuccessItems.Count)" -ForegroundColor Green
    Write-Host "  警告: $($Warnings.Count)" -ForegroundColor Yellow
    Write-Host "  错误: $($Issues.Count)" -ForegroundColor Red
    Write-Host ""
    
    if ($Issues.Count -eq 0) {
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  环境检查通过！可以开始构建。" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "构建命令:" -ForegroundColor White
        Write-Host "  .\build-standalone.bat" -ForegroundColor Cyan
        Write-Host ""
        return $true
    } else {
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "  发现 $($Issues.Count) 个问题需要修复" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host ""
        
        if ($Issues.Count -gt 0) {
            Write-Host "错误详情:" -ForegroundColor Red
            foreach ($issue in $Issues) {
                Write-Host "  - $issue" -ForegroundColor Red
            }
            Write-Host ""
        }
        
        if ($Warnings.Count -gt 0 -and $Verbose) {
            Write-Host "警告详情:" -ForegroundColor Yellow
            foreach ($warning in $Warnings) {
                Write-Host "  - $warning" -ForegroundColor Yellow
            }
            Write-Host ""
        }
        
        Write-Host "修复建议:" -ForegroundColor White
        Write-Host "  1. 安装 .NET Framework 4.8 Developer Pack" -ForegroundColor Cyan
        Write-Host "     https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Gray
        Write-Host "  2. 安装 Build Tools for Visual Studio 2022" -ForegroundColor Cyan
        Write-Host "     https://visualstudio.microsoft.com/downloads/" -ForegroundColor Gray
        Write-Host "  3. 运行诊断脚本并自动修复: .\diagnose-build.ps1 -Fix" -ForegroundColor Cyan
        Write-Host ""
        
        return $false
    }
}

# 主函数
function Main {
    Clear-Host
    Write-Host ""
    Write-Host "Word Translation Add-in 构建环境诊断工具" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Fix) {
        Write-Host "模式: 诊断并自动修复" -ForegroundColor Yellow
    } else {
        Write-Host "模式: 仅诊断" -ForegroundColor Yellow
    }
    Write-Host ""
    
    # 运行所有检查
    Test-DotNetFramework
    Test-MSBuild
    Test-NuGet
    Test-Office
    Test-VSTORuntime
    Test-ProjectFiles
    Test-DiskSpace
    
    # 显示报告
    $success = Show-DiagnosticReport
    
    if (-not $success) {
        exit 1
    }
}

# 执行主函数
Main
