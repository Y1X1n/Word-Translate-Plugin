namespace WordTranslationAddin.Models
{
    public class AppSettings
    {
        // 翻译引擎设置
        public TranslationEngineType PreferredEngine { get; set; }
        public string GoogleApiKey { get; set; }
        public string BaiduAppId { get; set; }
        public string BaiduSecretKey { get; set; }
        public string DeepLApiKey { get; set; }

        // 语言设置
        public string SourceLanguage { get; set; }
        public string TargetLanguage { get; set; }

        // 离线模式
        public bool EnableOfflineMode { get; set; }

        // 弹窗设置
        public int PopupWidth { get; set; }
        public int PopupHeight { get; set; }
        public PopupPosition PopupPosition { get; set; }
        public bool EnableAnimation { get; set; }

        // 功能设置
        public bool AutoCopyTranslation { get; set; }
        public bool ShowPhonetic { get; set; }
        public bool ShowExamples { get; set; }
        public int MaxHistoryItems { get; set; }

        // 外观设置
        public AppTheme Theme { get; set; }

        // 快捷键设置
        public string ShortcutKey { get; set; }

        // 自动翻译设置
        public bool EnableAutoTranslate { get; set; }
        public int MinTextLength { get; set; }
        public int MaxTextLength { get; set; }
    }

    public enum PopupPosition
    {
        Auto,
        TopLeft,
        TopRight,
        BottomLeft,
        BottomRight,
        Center,
        Cursor
    }

    public enum AppTheme
    {
        System,
        Light,
        Dark
    }
}
