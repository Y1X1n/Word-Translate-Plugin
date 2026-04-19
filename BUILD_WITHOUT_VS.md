# 无需 Visual Studio 构建 Word Translation Add-in 指南

本指南提供多种无需安装 Visual Studio 2022 即可构建项目的方案。

## 方案一：使用独立 MSBuild（推荐）

### 1. 安装必要工具

#### 1.1 下载并安装 .NET Framework 4.8 Developer Pack
- **下载链接**: https://dotnet.microsoft.com/download/dotnet-framework/net48
- **选择**: 
  - [.NET Framework 4.8 Developer Pack Offline Installer](https://go.microsoft.com/fwlink/?linkid=2088517)
- **安装步骤**:
  1. 下载 `ndp48-devpack-enu.exe`（约 60MB）
  2. 双击运行安装程序
  3. 按照向导完成安装
  4. 重启计算机

#### 1.2 下载并安装 Build Tools for Visual Studio 2022
- **下载链接**: https://visualstudio.microsoft.com/downloads/
- **选择**: 滚动到页面底部，找到 "Tools for Visual Studio" → "Build Tools for Visual Studio 2022"
- **安装步骤**:
  1. 下载 `vs_BuildTools.exe`
  2. 运行安装程序
  3. 在工作负载选项卡中，选择：
     - **.NET 桌面生成工具**
     - **Office/SharePoint 开发**
  4. 在单个组件选项卡中，确保选中：
     - **.NET Framework 4.8 目标包**
     - **.NET Framework 4.8 SDK**
     - **MSBuild**
  5. 点击安装（约 2-4GB）

### 2. 配置环境变量

安装完成后，将 MSBuild 添加到系统 PATH：

```powershell
# 以管理员身份运行 PowerShell
[Environment]::SetEnvironmentVariable(
    "Path",
    [Environment]::GetEnvironmentVariable("Path", "Machine") + ";C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin",
    "Machine"
)
```

或者手动配置：
1. 右键"此电脑" → 属性 → 高级系统设置
2. 环境变量 → 系统变量 → Path
3. 添加: `C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin`

### 3. 验证安装

```cmd
msbuild -version
```

应显示类似：
```
Microsoft (R) Build Engine version 17.x.x.x for .NET Framework
Copyright (C) Microsoft Corporation. All rights reserved.

17.x.x.x
```

### 4. 构建项目

#### 4.1 使用提供的构建脚本

```cmd
# 打开命令提示符或 PowerShell，切换到项目目录
cd "e:\trae projects\translation plugin"

# 运行构建脚本
build-standalone.bat
```

#### 4.2 手动构建

```cmd
# 还原 NuGet 包
nuget restore WordTranslationAddin.sln

# 构建项目
msbuild WordTranslationAddin.sln /p:Configuration=Release /p:Platform="Any CPU"
```

---

## 方案二：使用 .NET SDK + VS Code（轻量级）

### 1. 安装工具

#### 1.1 安装 .NET SDK 6.0 或更高版本
- **下载链接**: https://dotnet.microsoft.com/download
- **选择**: .NET SDK x64
- **安装**: 运行安装程序，按照向导完成

#### 1.2 安装 Visual Studio Code
- **下载链接**: https://code.visualstudio.com/
- **安装**: 运行安装程序

#### 1.3 安装 VS Code 扩展
1. 打开 VS Code
2. 按 `Ctrl+Shift+X` 打开扩展面板
3. 搜索并安装：
   - **C#** (由 Microsoft 提供)
   - **C# Dev Kit** (可选，提供更佳体验)

### 2. 构建项目

```bash
# 打开终端（在 VS Code 中按 Ctrl+`）

# 切换到项目目录
cd "WordTranslationAddin"

# 还原依赖
dotnet restore

# 构建项目
dotnet build --configuration Release
```

**注意**: 由于项目依赖 Office Interop 程序集，此方法可能需要额外配置。

---

## 方案三：使用 Rider（JetBrains）

### 1. 安装 JetBrains Rider
- **下载链接**: https://www.jetbrains.com/rider/
- **选择**: 社区版（免费）或专业版
- **安装**: 运行安装程序

### 2. 打开并构建项目
1. 打开 Rider
2. 选择 "Open" → 选择 `WordTranslationAddin.sln`
3. 等待项目加载和索引完成
4. 点击菜单 Build → Build Solution (Ctrl+F9)

---

## 方案四：使用 Docker 容器构建

### 1. 安装 Docker Desktop
- **下载链接**: https://www.docker.com/products/docker-desktop
- **安装**: 运行安装程序并启动 Docker

### 2. 使用预配置镜像构建

创建 `Dockerfile.build`:

```dockerfile
FROM mcr.microsoft.com/dotnet/framework/sdk:4.8-windowsservercore-ltsc2022

# 安装 MSBuild
RUN powershell -Command \
    Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vs_buildtools.exe' -OutFile 'vs_buildtools.exe'; \
    Start-Process -FilePath 'vs_buildtools.exe' -ArgumentList '--quiet', '--wait', '--add', 'Microsoft.VisualStudio.Workload.MSBuildTools', '--add', 'Microsoft.VisualStudio.Workload.NetCoreBuildTools' -NoNewWindow -Wait; \
    Remove-Item -Force 'vs_buildtools.exe'

WORKDIR /src
COPY . .

RUN nuget restore WordTranslationAddin.sln
RUN msbuild WordTranslationAddin.sln /p:Configuration=Release /p:Platform="Any CPU"
```

构建命令:
```powershell
docker build -f Dockerfile.build -t word-translation-build .
docker create --name extract word-translation-build
docker cp extract:/src/WordTranslationAddin/bin/Release ./Output
docker rm extract
```

---

## 方案五：使用 GitHub Actions 远程构建

### 1. 创建 GitHub 仓库
将代码推送到 GitHub 仓库。

### 2. 创建工作流文件

创建 `.github/workflows/build.yml`:

```yaml
name: Build Word Translation Add-in

on:
  push:
    branches: [ main, master ]
  pull_request:
    branches: [ main, master ]

jobs:
  build:
    runs-on: windows-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Setup MSBuild
      uses: microsoft/setup-msbuild@v1.1
      
    - name: Setup NuGet
      uses: NuGet/setup-nuget@v1.0.5
      
    - name: Restore NuGet packages
      run: nuget restore WordTranslationAddin.sln
      
    - name: Build
      run: msbuild WordTranslationAddin.sln /p:Configuration=Release /p:Platform="Any CPU"
      
    - name: Upload artifacts
      uses: actions/upload-artifact@v3
      with:
        name: WordTranslationAddin
        path: |
          WordTranslationAddin/bin/Release/*.dll
          WordTranslationAddin/bin/Release/*.config
```

### 3. 触发构建
推送代码到 GitHub，Actions 会自动构建并生成可下载的构建产物。

---

## 常见问题及解决方案

### 问题 1: "msbuild" 不是内部或外部命令

**原因**: MSBuild 未添加到系统 PATH

**解决方案**:
```powershell
# 查找 MSBuild 位置
Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Visual Studio" -Recurse -Filter "MSBuild.exe" -ErrorAction SilentlyContinue

# 将找到的路径添加到环境变量
```

### 问题 2: 找不到 .NET Framework 4.8 目标包

**原因**: 未安装 .NET Framework 4.8 Developer Pack

**解决方案**:
1. 下载并安装: https://dotnet.microsoft.com/download/dotnet-framework/net48
2. 或修改项目文件使用已安装的版本:
   ```xml
   <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
   ```

### 问题 3: NuGet 包还原失败

**原因**: 未安装 NuGet 或网络问题

**解决方案**:
```powershell
# 安装 NuGet
Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile "C:\Tools\nuget.exe"

# 添加到 PATH
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Tools", "Machine")
```

### 问题 4: Office Interop 程序集找不到

**原因**: 未安装 Office 或程序集不在 GAC 中

**解决方案**:
1. 安装 Microsoft Office
2. 或安装 Office Primary Interop Assemblies:
   - 下载: https://www.microsoft.com/en-us/download/details.aspx?id=54251
   - 或从 NuGet 安装: `Install-Package Microsoft.Office.Interop.Word`

### 问题 5: 构建成功但插件无法加载

**原因**: VSTO Runtime 未安装

**解决方案**:
下载并安装 VSTO Runtime:
- https://www.microsoft.com/en-us/download/details.aspx?id=54251

### 问题 6: WiX 安装程序构建失败

**原因**: 未安装 WiX Toolset

**解决方案**:
```powershell
# 使用 Chocolatey 安装
choco install wixtoolset

# 或手动下载安装
# https://wixtoolset.org/releases/
```

---

## 推荐的最低配置

| 组件 | 版本 | 大小 | 下载链接 |
|------|------|------|----------|
| .NET Framework 4.8 | 4.8 | 60MB | [下载](https://dotnet.microsoft.com/download/dotnet-framework/net48) |
| Build Tools 2022 | 17.x | 2-4GB | [下载](https://visualstudio.microsoft.com/downloads/) |
| NuGet CLI | 6.x | 5MB | [下载](https://www.nuget.org/downloads) |

**总磁盘空间需求**: 约 5-10GB

---

## 快速开始检查清单

- [ ] 安装 .NET Framework 4.8 Developer Pack
- [ ] 安装 Build Tools for Visual Studio 2022
- [ ] 配置 MSBuild 环境变量
- [ ] 安装 NuGet CLI
- [ ] 验证 `msbuild -version` 命令
- [ ] 运行 `build-standalone.bat`
- [ ] 检查输出目录中的 DLL 文件

---

## 获取帮助

如果遇到问题：
1. 查看 [TROUBLESHOOTING.md](TROUBLESHOOTING.md)
2. 运行诊断脚本: `powershell -File diagnose-build.ps1`
3. 提交 Issue 到项目仓库
