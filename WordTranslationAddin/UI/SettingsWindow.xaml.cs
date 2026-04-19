using System;
using System.Windows;
using System.Windows.Controls;
using WordTranslationAddin.Core;
using WordTranslationAddin.Models;
using WordTranslationAddin.Translation;

namespace WordTranslationAddin.UI
{
    public partial class SettingsWindow : Window
    {
        private readonly SettingsManager _settingsManager;

        public SettingsWindow(SettingsManager settingsManager)
        {
            _settingsManager = settingsManager ?? throw new ArgumentNullException(nameof(settingsManager));
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            var settings = _settingsManager.GetSettings();

            // 加载引擎选择
            foreach (ComboBoxItem item in EngineComboBox.Items)
            {
                if (Enum.TryParse<TranslationEngineType>(item.Tag.ToString(), out var engineType))
                {
                    if (engineType == settings.PreferredEngine)
                    {
                        EngineComboBox.SelectedItem = item;
                        break;
                    }
                }
            }

            // 加载语言选择
            foreach (ComboBoxItem item in LanguageComboBox.Items)
            {
                if (item.Tag.ToString() == settings.TargetLanguage)
                {
                    LanguageComboBox.SelectedItem = item;
                    break;
                }
            }

            // 加载API密钥
            GoogleApiKeyBox.Text = settings.GoogleApiKey ?? "";
            BaiduAppIdBox.Text = settings.BaiduAppId ?? "";
            BaiduSecretKeyBox.Text = settings.BaiduSecretKey ?? "";
            DeepLApiKeyBox.Text = settings.DeepLApiKey ?? "";

            // 加载功能设置
            OfflineModeCheckBox.IsChecked = settings.EnableOfflineMode;
            AutoCopyCheckBox.IsChecked = settings.AutoCopyTranslation;
            ShowPhoneticCheckBox.IsChecked = settings.ShowPhonetic;
            ShowExamplesCheckBox.IsChecked = settings.ShowExamples;
            EnableAnimationCheckBox.IsChecked = settings.EnableAnimation;

            // 加载弹窗位置
            foreach (ComboBoxItem item in PopupPositionComboBox.Items)
            {
                if (Enum.TryParse<PopupPosition>(item.Tag.ToString(), out var position))
                {
                    if (position == settings.PopupPosition)
                    {
                        PopupPositionComboBox.SelectedItem = item;
                        break;
                    }
                }
            }

            // 加载主题
            foreach (ComboBoxItem item in ThemeComboBox.Items)
            {
                if (Enum.TryParse<AppTheme>(item.Tag.ToString(), out var theme))
                {
                    if (theme == settings.Theme)
                    {
                        ThemeComboBox.SelectedItem = item;
                        break;
                    }
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _settingsManager.UpdateSettings(settings =>
                {
                    // 保存引擎选择
                    if (EngineComboBox.SelectedItem is ComboBoxItem engineItem && 
                        Enum.TryParse<TranslationEngineType>(engineItem.Tag.ToString(), out var engineType))
                    {
                        settings.PreferredEngine = engineType;
                    }

                    // 保存语言选择
                    if (LanguageComboBox.SelectedItem is ComboBoxItem langItem)
                    {
                        settings.TargetLanguage = langItem.Tag.ToString();
                    }

                    // 保存API密钥
                    settings.GoogleApiKey = GoogleApiKeyBox.Text?.Trim();
                    settings.BaiduAppId = BaiduAppIdBox.Text?.Trim();
                    settings.BaiduSecretKey = BaiduSecretKeyBox.Text?.Trim();
                    settings.DeepLApiKey = DeepLApiKeyBox.Text?.Trim();

                    // 保存功能设置
                    settings.EnableOfflineMode = OfflineModeCheckBox.IsChecked ?? true;
                    settings.AutoCopyTranslation = AutoCopyCheckBox.IsChecked ?? false;
                    settings.ShowPhonetic = ShowPhoneticCheckBox.IsChecked ?? true;
                    settings.ShowExamples = ShowExamplesCheckBox.IsChecked ?? true;
                    settings.EnableAnimation = EnableAnimationCheckBox.IsChecked ?? true;

                    // 保存弹窗位置
                    if (PopupPositionComboBox.SelectedItem is ComboBoxItem posItem &&
                        Enum.TryParse<PopupPosition>(posItem.Tag.ToString(), out var position))
                    {
                        settings.PopupPosition = position;
                    }

                    // 保存主题
                    if (ThemeComboBox.SelectedItem is ComboBoxItem themeItem &&
                        Enum.TryParse<AppTheme>(themeItem.Tag.ToString(), out var theme))
                    {
                        settings.Theme = theme;
                    }
                });

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
