# Word Translation Add-in 部署脚本
# 用于开发和测试环境的快速部署

param(
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Rebuild
)

$ErrorActionPreference = "Stop"

# 配置
$ProjectName = "WordTranslationAddin"
$AddInName = "WordTranslationAddin"
$SourceDir = ".\WordTranslationAddin\bin\Release"
$TargetDir = "$env:LOCALAPPDATA\TranslationPlugin\WordTranslationAddin"
$RegistryKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$AddInName"

function Write-Info($message) {
    Write-Host "[INFO] $message" -ForegroundColor Cyan
}

function Write-Success($message) {
    Write-Host "[SUCCESS] $message" -ForegroundColor Green
}

function Write-Error($message) {
    Write-Host "[ERROR] $message" -ForegroundColor Red
}

function Build-Project {
    Write-Info "开始构建项目..."
    
    $msbuild = & where.exe msbuild 2>$null
    if (-not $msbuild) {
        throw "找不到 MSBuild。请确保已安装 Visual Studio。"
    }
    
    & msbuild "$ProjectName.sln" /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
    
    if ($LASTEXITCODE -ne 0) {
        throw "构建失败！"
    }
    
    Write-Success "项目构建成功"
}

function Install-AddIn {
    Write-Info "开始安装插件..."
    
    # 创建目标目录
    if (-not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        Write-Info "创建目录: $TargetDir"
    }
    
    # 复制文件
    $files = @(
        "$SourceDir\$AddInName.dll",
        "$SourceDir\$AddInName.dll.config",
        "$SourceDir\Newtonsoft.Json.dll"
    )
    
    foreach ($file in $files) {
        if (Test-Path $file) {
            Copy-Item $file $TargetDir -Force
            Write-Info "复制文件: $(Split-Path $file -Leaf)"
        } else {
            Write-Error "文件不存在: $file"
        }
    }
    
    # 注册插件
    if (-not (Test-Path $RegistryKey)) {
        New-Item -Path $RegistryKey -Force | Out-Null
    }
    
    $manifestPath = "$TargetDir\$AddInName.dll"
    Set-ItemProperty -Path $RegistryKey -Name "Description" -Value "Word Translation Add-in"
    Set-ItemProperty -Path $RegistryKey -Name "FriendlyName" -Value "Word Translation Add-in"
    Set-ItemProperty -Path $RegistryKey -Name "Manifest" -Value "$manifestPath|vsto|vstolocal"
    Set-ItemProperty -Path $RegistryKey -Name "LoadBehavior" -Value 3 -Type DWord
    
    Write-Success "插件安装成功"
    Write-Info "安装位置: $TargetDir"
}

function Uninstall-AddIn {
    Write-Info "开始卸载插件..."
    
    # 删除注册表项
    if (Test-Path $RegistryKey) {
        Remove-Item -Path $RegistryKey -Recurse -Force
        Write-Info "删除注册表项"
    }
    
    # 删除文件
    if (Test-Path $TargetDir) {
        Remove-Item -Path $TargetDir -Recurse -Force
        Write-Info "删除文件: $TargetDir"
    }
    
    Write-Success "插件卸载成功"
}

function Show-Help {
    Write-Host @"
Word Translation Add-in 部署脚本

用法:
    .\deploy.ps1 [-Install] [-Uninstall] [-Rebuild]

参数:
    -Install     安装插件
    -Uninstall   卸载插件
    -Rebuild     重新构建项目

示例:
    .\deploy.ps1 -Install          # 构建并安装
    .\deploy.ps1 -Install -Rebuild # 重新构建并安装
    .\deploy.ps1 -Uninstall        # 卸载插件
"@
}

# 主逻辑
if ($Uninstall) {
    Uninstall-AddIn
    exit 0
}

if ($Rebuild) {
    Build-Project
}

if ($Install) {
    if (-not $Rebuild) {
        # 检查是否存在构建输出
        if (-not (Test-Path "$SourceDir\$AddInName.dll")) {
            Write-Info "未找到构建输出，开始构建..."
            Build-Project
        }
    }
    
    Install-AddIn
    
    Write-Host ""
    Write-Success "部署完成！请重启 Microsoft Word 以使用插件。"
} else {
    Show-Help
}
