using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WordTranslationAddin.Core;
using WordTranslationAddin.Models;
using WordTranslationAddin.Translation;

namespace WordTranslationAddin.UI
{
    public partial class HistoryWindow : Window
    {
        private readonly HistoryManager _historyManager;
        private List<TranslationHistory> _currentHistory;

        public HistoryWindow(HistoryManager historyManager)
        {
            _historyManager = historyManager ?? throw new ArgumentNullException(nameof(historyManager));
            InitializeComponent();
            LoadHistory();
        }

        private void LoadHistory()
        {
            _currentHistory = _historyManager.GetHistory(100);
            RefreshHistoryList();
        }

        private void RefreshHistoryList()
        {
            HistoryListPanel.Children.Clear();
            
            if (_currentHistory.Count == 0)
            {
                EmptyStatePanel.Visibility = Visibility.Visible;
                CountTextBlock.Text = "共 0 条记录";
                return;
            }

            EmptyStatePanel.Visibility = Visibility.Collapsed;
            CountTextBlock.Text = $"共 {_currentHistory.Count} 条记录";

            foreach (var item in _currentHistory)
            {
                var historyItem = CreateHistoryItem(item);
                HistoryListPanel.Children.Add(historyItem);
            }
        }

        private Border CreateHistoryItem(TranslationHistory item)
        {
            var border = new Border
            {
                Style = FindResource("HistoryItemStyle") as Style,
                Background = Brushes.White,
                Cursor = Cursors.Hand
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // 内容区
            var contentStack = new StackPanel();
            
            // 源文本
            var sourceText = new TextBlock
            {
                Text = item.SourceText,
                FontSize = 14,
                FontWeight = FontWeights.Medium,
                Foreground = FindResource("TextPrimaryColor") as Brush,
                TextTrimming = TextTrimming.CharacterEllipsis,
                MaxWidth = 500
            };
            contentStack.Children.Add(sourceText);

            // 翻译结果
            var translatedText = new TextBlock
            {
                Text = item.TranslatedText,
                FontSize = 13,
                Foreground = FindResource("PrimaryColor") as Brush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 0),
                MaxWidth = 500
            };
            contentStack.Children.Add(translatedText);

            // 元信息
            var metaStack = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 0) };
            
            var engineText = new TextBlock
            {
                Text = GetEngineDisplayName(item.EngineType),
                FontSize = 11,
                Foreground = FindResource("TextTertiaryColor") as Brush
            };
            metaStack.Children.Add(engineText);

            var separator = new TextBlock
            {
                Text = " • ",
                FontSize = 11,
                Foreground = FindResource("TextTertiaryColor") as Brush,
                Margin = new Thickness(4, 0)
            };
            metaStack.Children.Add(separator);

            var timeText = new TextBlock
            {
                Text = item.Timestamp.ToString("yyyy-MM-dd HH:mm"),
                FontSize = 11,
                Foreground = FindResource("TextTertiaryColor") as Brush
            };
            metaStack.Children.Add(timeText);

            contentStack.Children.Add(metaStack);
            Grid.SetColumn(contentStack, 0);
            grid.Children.Add(contentStack);

            // 操作按钮
            var actionStack = new StackPanel { Orientation = Orientation.Horizontal };
            
            var copyButton = new Button
            {
                Style = FindResource("IconButtonStyle") as Style,
                ToolTip = "复制翻译结果",
                Tag = item
            };
            var copyIcon = new Path
            {
                Data = Geometry.Parse("M19,21H8V7H19M19,5H8A2,2 0 0,0 6,7V21A2,2 0 0,0 8,23H19A2,2 0 0,0 21,21V7A2,2 0 0,0 19,5M16,1H4A2,2 0 0,0 2,3V17H4V3H16V1Z"),
                Fill = FindResource("TextSecondaryColor") as Brush,
                Width = 16,
                Height = 16,
                Stretch = Stretch.Uniform
            };
            copyButton.Content = copyIcon;
            copyButton.Click += CopyButton_Click;
            actionStack.Children.Add(copyButton);

            var deleteButton = new Button
            {
                Style = FindResource("IconButtonStyle") as Style,
                ToolTip = "删除",
                Tag = item,
                Margin = new Thickness(4, 0, 0, 0)
            };
            var deleteIcon = new Path
            {
                Data = Geometry.Parse("M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z"),
                Fill = FindResource("TextSecondaryColor") as Brush,
                Width = 16,
                Height = 16,
                Stretch = Stretch.Uniform
            };
            deleteButton.Content = deleteIcon;
            deleteButton.Click += DeleteButton_Click;
            actionStack.Children.Add(deleteButton);

            Grid.SetColumn(actionStack, 1);
            grid.Children.Add(actionStack);

            border.Child = grid;

            // 点击复制
            border.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ClickCount == 2)
                {
                    Clipboard.SetText(item.TranslatedText);
                    StatusTextBlock.Text = "已复制到剪贴板";
                }
            };

            return border;
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

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string keyword = SearchBox.Text?.Trim();
            
            if (string.IsNullOrEmpty(keyword))
            {
                _currentHistory = _historyManager.GetHistory(100);
            }
            else
            {
                _currentHistory = _historyManager.SearchHistory(keyword);
            }
            
            RefreshHistoryList();
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is TranslationHistory item)
            {
                Clipboard.SetText(item.TranslatedText);
                StatusTextBlock.Text = "已复制到剪贴板";
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is TranslationHistory item)
            {
                var result = MessageBox.Show("确定要删除这条历史记录吗？", "确认删除", 
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    _historyManager.RemoveFromHistory(item.Id);
                    LoadHistory();
                    StatusTextBlock.Text = "已删除";
                }
            }
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("确定要清空所有历史记录吗？此操作不可恢复。", "确认清空", 
                MessageBoxButton.YesNo, MessageBoxImage.Warning);
            
            if (result == MessageBoxResult.Yes)
            {
                _historyManager.ClearHistory();
                LoadHistory();
                StatusTextBlock.Text = "历史记录已清空";
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
