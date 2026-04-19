# Word Translation Add-in

Microsoft Office 2024 Word 专用翻译插件

## 功能特性

- **智能文本识别**：自动识别用户框选的英文单词或句子
- **多翻译引擎支持**：集成谷歌翻译、百度翻译和DeepL翻译
- **精美UI界面**：采用Microsoft Office Fluent Design设计语言
- **离线翻译模式**：内置常用词库，无需网络即可翻译
- **翻译历史记录**：保存和查看翻译历史
- **自定义设置**：支持弹窗大小、位置、样式自定义

## 系统要求

- Microsoft Office 2024 (Word)
- Windows 10/11 (32位或64位)
- .NET Framework 4.8 或更高版本
- 网络连接（在线翻译功能）

## 安装方法

### 方法1：使用安装程序
1. 下载 `WordTranslationAddin.msi`
2. 双击运行安装程序
3. 按照向导完成安装
4. 重启 Microsoft Word

### 方法2：手动安装
1. 构建项目生成 DLL 文件
2. 将文件复制到 `%LocalAppData%\TranslationPlugin\WordTranslationAddin\`
3. 注册 COM 组件
4. 重启 Microsoft Word

## 使用方法

1. 在 Word 文档中框选要翻译的英文文本
2. 插件会自动识别并弹出翻译结果窗口
3. 点击设置按钮配置翻译引擎和API密钥
4. 点击历史记录按钮查看翻译历史

## API 配置

### 谷歌翻译
1. 访问 [Google Cloud Console](https://console.cloud.google.com/)
2. 创建项目并启用 Cloud Translation API
3. 创建 API 密钥
4. 在插件设置中输入 API 密钥

### 百度翻译
1. 访问 [百度翻译开放平台](https://fanyi-api.baidu.com/)
2. 注册并创建应用
3. 获取 App ID 和密钥
4. 在插件设置中输入相应信息

### DeepL
1. 访问 [DeepL API](https://www.deepl.com/pro-api)
2. 注册并获取 API 密钥
3. 在插件设置中输入 API 密钥

## 项目结构

```
WordTranslationAddin/
├── Core/                       # 核心功能模块
│   ├── TextSelectionMonitor.cs # 文本选择监控
│   ├── TranslationManager.cs   # 翻译管理器
│   ├── SettingsManager.cs      # 设置管理器
│   └── HistoryManager.cs       # 历史记录管理器
├── Translation/                # 翻译引擎
│   ├── ITranslationEngine.cs   # 翻译引擎接口
│   ├── GoogleTranslateEngine.cs    # 谷歌翻译
│   ├── BaiduTranslateEngine.cs     # 百度翻译
│   ├── DeepLTranslateEngine.cs     # DeepL翻译
│   └── OfflineDictionaryEngine.cs  # 离线词典
├── UI/                         # 用户界面
│   ├── TranslationPopup.xaml   # 翻译弹窗
│   ├── SettingsWindow.xaml     # 设置窗口
│   └── HistoryWindow.xaml      # 历史记录窗口
├── Models/                     # 数据模型
│   ├── TranslationResult.cs    # 翻译结果
│   ├── TranslationHistory.cs   # 翻译历史
│   └── AppSettings.cs          # 应用设置
├── Helpers/                    # 辅助类
│   ├── WindowPositionHelper.cs # 窗口位置帮助
│   └── TextHelper.cs           # 文本处理帮助
└── Resources/                  # 资源文件
    └── OfflineDictionary.json  # 离线词典数据
```

## 构建项目

### 前提条件
- Visual Studio 2022
- .NET Framework 4.8
- Office Developer Tools
- WiX Toolset (用于创建安装程序)

### 构建步骤
1. 打开 `WordTranslationAddin.sln`
2. 选择 Release 配置
3. 生成解决方案
4. 运行 Installer 项目生成 MSI 安装包

## 性能指标

- 翻译响应时间：平均 < 2秒
- 内存占用：峰值 < 100MB
- CPU使用率：翻译过程中 < 20%
- 支持文本长度：单次最多 500 字符

## 许可证

MIT License

## 技术支持

如有问题或建议，请提交 Issue 或联系开发团队。
