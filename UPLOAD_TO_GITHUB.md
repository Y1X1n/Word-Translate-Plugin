# 手动上传代码到 GitHub 指南

由于网络限制，你需要手动上传代码到 GitHub 仓库来触发自动构建。

## 步骤 1：准备代码文件

代码已经准备好在以下目录：
```
e:\trae projects\translation plugin\
```

## 步骤 2：上传到 GitHub

### 方法 A：通过 GitHub 网页上传（推荐）

1. **访问你的仓库**
   - 打开浏览器访问：https://github.com/Y1X1n/Word-Translate-Plugin

2. **上传文件**
   - 点击页面中间的 "uploading an existing file" 链接
   - 或者点击 "Add file" → "Upload files"

3. **选择文件**
   - 将所有项目文件拖拽到上传区域，或点击 "choose your files" 选择文件
   - 需要上传的文件包括：
     - `.github/workflows/build.yml` (GitHub Actions 配置)
     - `WordTranslationAddin.sln` (解决方案文件)
     - `WordTranslationAddin/` 文件夹及其所有内容
     - `README.md`
     - 其他所有文件

4. **提交更改**
   - 在 "Commit changes" 部分，输入提交信息：
     ```
     Initial commit: Word Translation Add-in project
     ```
   - 点击 "Commit changes"

### 方法 B：使用 GitHub Desktop

1. **下载 GitHub Desktop**
   - 访问：https://desktop.github.com/
   - 下载并安装

2. **克隆仓库**
   - 打开 GitHub Desktop
   - 点击 "File" → "Clone repository"
   - 选择 "URL" 标签
   - 输入：`https://github.com/Y1X1n/Word-Translate-Plugin.git`
   - 选择本地路径，点击 "Clone"

3. **复制文件**
   - 将 `e:\trae projects\translation plugin\` 中的所有文件
   - 复制到克隆的仓库文件夹中

4. **提交并推送**
   - 在 GitHub Desktop 中，输入提交信息
   - 点击 "Commit to main"
   - 点击 "Push origin"

## 步骤 3：查看自动构建

上传完成后，GitHub Actions 会自动开始构建：

1. **查看构建状态**
   - 访问：https://github.com/Y1X1n/Word-Translate-Plugin/actions
   - 你会看到正在运行的 "Build Word Translation Add-in" 工作流

2. **等待构建完成**
   - 构建通常需要 3-5 分钟
   - 绿色 ✓ 表示构建成功
   - 红色 ✗ 表示构建失败

3. **下载构建产物**
   - 构建完成后，点击最新的工作流运行记录
   - 滚动到页面底部的 "Artifacts" 部分
   - 下载以下文件：
     - `WordTranslationAddin-Release-vX` - 包含 DLL 文件的压缩包
     - `WordTranslationAddin-Installer-vX` - MSI 安装程序（如果构建成功）

## 构建产物说明

下载的文件包含：

| 文件 | 说明 |
|------|------|
| `WordTranslationAddin.dll` | 插件主程序 |
| `Newtonsoft.Json.dll` | JSON 解析库 |
| `WordTranslationAddin.dll.config` | 配置文件 |
| `WordTranslationAddin.msi` | 安装程序（可选）|

## 安装插件

### 方法 1：使用 MSI 安装程序（如果有）
1. 下载 `WordTranslationAddin.msi`
2. 双击运行安装程序
3. 按照向导完成安装
4. 重启 Microsoft Word

### 方法 2：手动安装
1. 解压下载的文件
2. 复制到 `%LocalAppData%\TranslationPlugin\WordTranslationAddin\`
3. 运行 `deploy.ps1` 脚本注册插件
4. 重启 Microsoft Word

## 验证安装

1. 打开 Microsoft Word
2. 点击 "文件" → "选项" → "加载项"
3. 在 "管理" 下拉框中选择 "COM 加载项"
4. 点击 "转到"
5. 确认 "Word Translation Add-in" 在列表中且已勾选

## 故障排除

如果构建失败：
1. 点击失败的构建记录查看日志
2. 检查错误信息
3. 参考 `TROUBLESHOOTING.md` 文件

如果需要帮助：
1. 查看构建日志中的具体错误
2. 检查是否所有文件都已正确上传
3. 确认 `.github/workflows/build.yml` 文件存在

## 下次更新代码

当你修改代码后：
1. 在 GitHub 网页上点击 "Add file" → "Upload files"
2. 上传修改后的文件
3. 输入提交信息，如："Update translation engine"
4. 点击 "Commit changes"
5. GitHub Actions 会自动重新构建
