# 构建问题排查指南

本文档整理了构建过程中可能遇到的常见问题及解决方案。

## 目录
1. [环境相关问题](#环境相关问题)
2. [构建错误](#构建错误)
3. [运行时问题](#运行时问题)
4. [安装问题](#安装问题)

---

## 环境相关问题

### 问题：找不到 MSBuild

**症状**：
```
'msbuild' 不是内部或外部命令，也不是可运行的程序
```

**解决方案**：

1. **安装 Build Tools for Visual Studio 2022**
   ```powershell
   # 下载并安装
   Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vs_buildtools.exe" -OutFile "vs_buildtools.exe"
   Start-Process -FilePath "vs_buildtools.exe" -ArgumentList "--quiet", "--wait", "--add", "Microsoft.VisualStudio.Workload.MSBuildTools", "--add", "Microsoft.VisualStudio.Workload.NetCoreBuildTools" -Wait
   ```

2. **手动添加到 PATH**
   ```powershell
   [Environment]::SetEnvironmentVariable(
       "Path",
       [Environment]::GetEnvironmentVariable("Path", "Machine") + ";C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin",
       "Machine"
   )
   ```

3. **验证安装**
   ```cmd
   msbuild -version
   ```

---

### 问题：缺少 .NET Framework 4.8

**症状**：
```
error MSB3644: 找不到 .NET Framework v4.8 的引用程序集
```

**解决方案**：

1. **下载并安装 Developer Pack**
   - 链接：https://dotnet.microsoft.com/download/dotnet-framework/net48
   - 选择：Developer Pack Offline Installer

2. **使用离线安装包**
   ```powershell
   # 如果无法联网，使用离线安装包
   ndp48-devpack-enu.exe /quiet /norestart
   ```

3. **验证安装**
   ```powershell
   Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' | Get-ItemPropertyValue -Name Release
   # 应返回 528040 或更高
   ```

---

### 问题：NuGet 包还原失败

**症状**：
```
Unable to find version 'x.x.x' of package 'Newtonsoft.Json'
```

**解决方案**：

1. **清除 NuGet 缓存**
   ```cmd
   nuget locals all -clear
   ```

2. **手动还原包**
   ```cmd
   nuget restore WordTranslationAddin.sln -Source https://api.nuget.org/v3/index.json
   ```

3. **检查网络连接**
   ```powershell
   Test-NetConnection -ComputerName api.nuget.org -Port 443
   ```

4. **配置代理（如需要）**
   ```cmd
   nuget config -set http_proxy=http://proxy.company.com:8080
   ```

---

## 构建错误

### 问题：Office Interop 程序集找不到

**症状**：
```
error CS0234: The type or namespace name 'Office' does not exist in the namespace 'Microsoft'
```

**解决方案**：

1. **安装 Office Primary Interop Assemblies**
   ```powershell
   # 下载 Office PIA
   # https://www.microsoft.com/en-us/download/details.aspx?id=54251
   ```

2. **使用 NuGet 包替代**
   ```xml
   <!-- 在项目文件中添加 -->
   <PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.4797.1004" />
   ```

3. **手动添加引用**
   - 找到 `Microsoft.Office.Interop.Word.dll`
   - 通常位于：`C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word`

---

### 问题：VSTO 相关错误

**症状**：
```
error MSB3482: SignTool reported an error: SignTool Error
```

**解决方案**：

1. **禁用强名称验证（开发环境）**
   ```cmd
   sn -Vr *,*
   ```

2. **创建测试证书**
   ```powershell
   $cert = New-SelfSignedCertificate -DnsName "WordTranslationAddin" -CertStoreLocation "cert:\LocalMachine\My" -Type CodeSigning
   Export-Certificate -Cert $cert -FilePath "WordTranslationAddin.cer"
   ```

3. **修改项目文件跳过签名**
   ```xml
   <PropertyGroup>
     <SignAssembly>false</SignAssembly>
     <DelaySign>false</DelaySign>
   </PropertyGroup>
   ```

---

### 问题：编译器版本不匹配

**症状**：
```
error MSB8020: The build tools for Visual Studio 2019 cannot be found
```

**解决方案**：

1. **升级项目文件**
   ```xml
   <!-- 修改 WordTranslationAddin.csproj -->
   <PlatformToolset>v143</PlatformToolset>  <!-- VS 2022 -->
   ```

2. **指定工具集版本**
   ```cmd
   msbuild WordTranslationAddin.sln /p:PlatformToolset=v143
   ```

---

## 运行时问题

### 问题：插件无法加载到 Word

**症状**：
- Word 启动时插件未显示
- 插件在 COM 加载项列表中显示为"已卸载"

**解决方案**：

1. **检查注册表项**
   ```powershell
   # 检查插件注册
   Get-ItemProperty "HKCU:\Software\Microsoft\Office\Word\Addins\WordTranslationAddin"
   
   # 确保 LoadBehavior 为 3
   Set-ItemProperty "HKCU:\Software\Microsoft\Office\Word\Addins\WordTranslationAddin" -Name "LoadBehavior" -Value 3
   ```

2. **检查 VSTO Runtime**
   ```powershell
   # 验证 VSTO Runtime 安装
   Test-Path "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
   
   # 下载链接
   # https://www.microsoft.com/en-us/download/details.aspx?id=54251
   ```

3. **启用插件加载日志**
   ```powershell
   # 启用 VSTO 日志
   Set-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Word\Options" -Name "VstoLog" -Value 1
   ```

4. **检查事件查看器**
   ```powershell
   # 查看应用程序日志
   Get-EventLog -LogName Application -Source "VSTO 4.0" -Newest 10
   ```

---

### 问题：翻译功能不工作

**症状**：
- 选择文本后无弹窗出现
- 弹窗显示但翻译失败

**解决方案**：

1. **检查 API 配置**
   ```powershell
   # 检查设置文件
   $settingsPath = "$env:LOCALAPPDATA\WordTranslationAddin\settings.json"
   Get-Content $settingsPath | ConvertFrom-Json
   ```

2. **测试网络连接**
   ```powershell
   # 测试翻译 API
   Test-NetConnection -ComputerName translate.googleapis.com -Port 443
   Test-NetConnection -ComputerName fanyi-api.baidu.com -Port 443
   Test-NetConnection -ComputerName api-free.deepl.com -Port 443
   ```

3. **检查防火墙设置**
   ```powershell
   # 查看防火墙规则
   Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*Word*" }
   ```

---

## 安装问题

### 问题：MSI 安装程序无法运行

**症状**：
- 双击 MSI 文件无反应
- 安装过程中出现错误代码

**解决方案**：

1. **检查 Windows Installer 服务**
   ```cmd
   sc query msiserver
   net start msiserver
   ```

2. **启用 MSI 日志**
   ```cmd
   msiexec /i WordTranslationAddin.msi /l*v install.log
   ```

3. **检查系统要求**
   ```powershell
   # 检查 .NET Framework
   Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Release
   
   # 检查 Office
   Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
   ```

---

### 问题：卸载后无法重新安装

**症状**：
- 提示"已安装更新版本"
- 卸载后注册表残留

**解决方案**：

1. **手动清理注册表**
   ```powershell
   # 删除插件注册表项
   Remove-Item -Path "HKCU:\Software\Microsoft\Office\Word\Addins\WordTranslationAddin" -Recurse -Force -ErrorAction SilentlyContinue
   
   # 删除安装目录
   Remove-Item -Path "$env:LOCALAPPDATA\TranslationPlugin" -Recurse -Force -ErrorAction SilentlyContinue
   ```

2. **使用 MSI 清理工具**
   ```powershell
   # 查找产品代码
   Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Word Translation*" }
   
   # 强制卸载
   msiexec /x {PRODUCT-CODE-HERE} /qn
   ```

---

## 诊断工具

### 运行诊断脚本

```powershell
# 全面诊断
.\diagnose-build.ps1

# 诊断并自动修复
.\diagnose-build.ps1 -Fix

# 详细诊断
.\diagnose-build.ps1 -Verbose
```

### 收集诊断信息

```powershell
# 创建诊断报告
$report = @{
    Date = Get-Date
    OS = (Get-CimInstance Win32_OperatingSystem).Caption
    DotNet = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Version).Version
    MSBuild = (Get-Command msbuild -ErrorAction SilentlyContinue).Version
    Office = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe" -ErrorAction SilentlyContinue).Version
}

$report | ConvertTo-Json | Out-File "diagnostic-report.json"
```

---

## 获取帮助

如果以上方案无法解决问题：

1. **查看详细日志**
   - 构建日志：`build-standalone.log`
   - 安装日志：`%TEMP%\MSI*.log`
   - VSTO 日志：`%TEMP%\VSTO*.log`

2. **提交 Issue**
   - 提供操作系统版本
   - 提供 Office 版本
   - 附上完整错误信息
   - 附上诊断报告：`diagnose-build.ps1 -Verbose`

3. **社区支持**
   - Stack Overflow: [vsto] 标签
   - Microsoft Q&A
   - GitHub Discussions
