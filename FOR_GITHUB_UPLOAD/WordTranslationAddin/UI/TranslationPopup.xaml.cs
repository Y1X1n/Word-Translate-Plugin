using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WordTranslationAddin.Core;
using WordTranslationAddin.Models;
using WordTranslationAddin.Translation;

namespace WordTranslationAddin.UI
{
    public partial class TranslationPopup : Window
    {
        private readonly SettingsManager _settingsManager;
        private readonly HistoryManager _historyManager;
        private TranslationResult _currentResult;
        private string _sourceText;

        public TranslationPopup(SettingsManager settingsManager, HistoryManager historyManager)
        {
            _settingsManager = settingsManager ?? throw new ArgumentNullException(nameof(settingsManager));
            _historyManager = historyManager ?? throw new ArgumentNullException(nameof(historyManager));
            
            InitializeComponent();
            InitializeEngineButtons();
        }

        private void InitializeEngineButtons()
        {
            var settings = _settingsManager.GetSettings();
            
            switch (settings.PreferredEngine)
            {
                case TranslationEngineType.Google:
                    GoogleEngineButton.IsChecked = true;
                    break;
                case TranslationEngineType.Baidu:
                    BaiduEngineButton.IsChecked = true;
                    break;
                case TranslationEngineType.DeepL:
                    DeepLEngineButton.IsChecked = true;
                    break;
            }
        }

        public void SetSourceText(string text)
        {
            _sourceText = text;
            SourceTextBlock.Text = text;
            StatusTextBlock.Text = "正在翻译...";
            
            // 清空之前的结果
            TranslatedTextBlock.Text = "";
            PhoneticBlock.Visibility = Visibility.Collapsed;
            PartOfSpeechBlock.Visibility = Visibility.Collapsed;
            DefinitionBlock.Visibility = Visibility.Collapsed;
            ExamplesPanel.Visibility = Visibility.Collapsed;
            OfflineBadge.Visibility = Visibility.Collapsed;
            ExamplesPanel.Children.Clear();
            
            // 重新添加例句标题
            var examplesTitle = new TextBlock
            {
                Text = "例句",
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = FindResource("TextSecondaryColor") as Brush,
                Margin = new Thickness(0, 0, 0, 8)
            };
            ExamplesPanel.Children.Add(examplesTitle);
        }

        public void SetTranslationResult(TranslationResult result)
        {
            _currentResult = result;
            
            if (result.IsError)
            {
                TranslatedTextBlock.Text = result.TranslatedText;
                TranslatedTextBlock.Foreground = Brushes.Red;
                StatusTextBlock.Text = "翻译失败";
                return;
            }

            TranslatedTextBlock.Text = result.TranslatedText;
            TranslatedTextBlock.Foreground = FindResource("PrimaryColor") as Brush;
            
            // 显示引擎信息
            string engineName = GetEngineDisplayName(result.EngineType);
            if (result.IsOffline)
            {
                engineName += " (离线)";
                OfflineBadge.Visibility = Visibility.Visible;
            }
            else if (result.IsFallback)
            {
                engineName += " (备用)";
            }
            EngineTextBlock.Text = engineName;
            
            // 显示状态
            StatusTextBlock.Text = $"耗时: {result.ResponseTimeMs}ms";
            
            // 显示单词详细信息
            if (result.IsWord)
            {
                if (!string.IsNullOrEmpty(result.Phonetic))
                {
                    PhoneticBlock.Text = result.Phonetic;
                    PhoneticBlock.Visibility = Visibility.Visible;
                }

                if (!string.IsNullOrEmpty(result.PartOfSpeech))
                {
                    PartOfSpeechBlock.Text = result.PartOfSpeech;
                    PartOfSpeechBlock.Visibility = Visibility.Visible;
                }

                if (!string.IsNullOrEmpty(result.Definition))
                {
                    DefinitionBlock.Text = result.Definition;
                    DefinitionBlock.Visibility = Visibility.Visible;
                }

                if (result.Examples != null && result.Examples.Count > 0)
                {
                    foreach (var example in result.Examples)
                    {
                        var exampleText = new TextBlock
                        {
                            Text = "• " + example,
                            TextWrapping = TextWrapping.Wrap,
                            FontSize = 12,
                            Foreground = FindResource("TextSecondaryColor") as Brush,
                            Margin = new Thickness(0, 0, 0, 4)
                        };
                        ExamplesPanel.Children.Add(exampleText);
                    }
                    ExamplesPanel.Visibility = Visibility.Visible;
                }
            }

            // 保存到历史记录
            _historyManager.AddToHistory(result);
        }

        public void SetPosition(System.Drawing.Point screenPosition)
        {
            var settings = _settingsManager.GetSettings();
            
            double left = screenPosition.X;
            double top = screenPosition.Y;

            // 确保窗口不超出屏幕边界
            var screen = SystemParameters.WorkArea;
            
            if (left + Width > screen.Width)
            {
                left = screenPosition.X - Width;
            }
            
            if (top + Height > screen.Height)
            {
                top = screenPosition.Y - Height;
            }

            Left = Math.Max(screen.Left, left);
            Top = Math.Max(screen.Top, top);
        }

        private string GetEngineDisplayName(TranslationEngineType engineType)
        {
            switch (engineType)
            {
                case TranslationEngineType.Google:
                    return "谷歌翻译";
                case TranslationEngineType.Baidu:
                    return "百度翻译";
                case TranslationEngineType.DeepL:
                    return "DeepL";
                case TranslationEngineType.Offline:
                    return "离线词典";
                default:
                    return "未知引擎";
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentResult != null && !string.IsNullOrEmpty(_currentResult.TranslatedText))
            {
                Clipboard.SetText(_currentResult.TranslatedText);
                StatusTextBlock.Text = "已复制到剪贴板";
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow(_settingsManager);
            settingsWindow.Owner = this;
            settingsWindow.ShowDialog();
        }

        private void HistoryButton_Click(object sender, RoutedEventArgs e)
        {
            var historyWindow = new HistoryWindow(_historyManager);
            historyWindow.Owner = this;
            historyWindow.ShowDialog();
        }

        private void EngineButton_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton radioButton && radioButton.IsChecked == true)
            {
                TranslationEngineType engineType;
                
                if (radioButton == GoogleEngineButton)
                    engineType = TranslationEngineType.Google;
                else if (radioButton == BaiduEngineButton)
                    engineType = TranslationEngineType.Baidu;
                else if (radioButton == DeepLEngineButton)
                    engineType = TranslationEngineType.DeepL;
                else
                    return;

                _settingsManager.UpdateSettings(s => s.PreferredEngine = engineType);
                
                // 重新翻译
                if (!string.IsNullOrEmpty(_sourceText))
                {
                    // 这里可以触发重新翻译
                }
            }
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
            // 点击外部时关闭窗口
            // Close();
        }
    }
}
