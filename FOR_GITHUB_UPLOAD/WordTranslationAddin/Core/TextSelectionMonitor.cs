using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordTranslationAddin.Core
{
    public class TextSelectionMonitor
    {
        private readonly Word.Application _application;
        private Timer _debounceTimer;
        private readonly object _lockObject = new object();
        private string _lastSelectedText = string.Empty;
        private DateTime _lastSelectionTime = DateTime.MinValue;
        private readonly TimeSpan _debounceInterval = TimeSpan.FromMilliseconds(300);
        private readonly TimeSpan _minSelectionInterval = TimeSpan.FromMilliseconds(500);

        public event EventHandler<TextSelectedEventArgs> TextSelected;

        public TextSelectionMonitor(Word.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void StartMonitoring()
        {
            // 监控已在Application事件中处理
        }

        public void StopMonitoring()
        {
            _debounceTimer?.Dispose();
        }

        public void HandleSelectionChange(Selection selection)
        {
            if (selection == null || selection.Range == null)
                return;

            string selectedText = selection.Range.Text?.Trim();
            
            if (string.IsNullOrWhiteSpace(selectedText))
                return;

            if (selectedText.Length > 500)
                return;

            if (selectedText == _lastSelectedText)
                return;

            var now = DateTime.Now;
            if ((now - _lastSelectionTime) < _minSelectionInterval)
                return;

            lock (_lockObject)
            {
                _debounceTimer?.Dispose();
                _debounceTimer = new Timer(DebounceCallback, selectedText, 
                    (int)_debounceInterval.TotalMilliseconds, Timeout.Infinite);
            }
        }

        private void DebounceCallback(object state)
        {
            string selectedText = state as string;
            if (string.IsNullOrEmpty(selectedText))
                return;

            try
            {
                _application.ActiveWindow.GetPoint(out int left, out int top, 
                    out int width, out int height, _application.Selection.Range);
                
                var screenPosition = new Point(left + width, top);
                
                _lastSelectedText = selectedText;
                _lastSelectionTime = DateTime.Now;

                TextSelected?.Invoke(this, new TextSelectedEventArgs
                {
                    SelectedText = selectedText,
                    ScreenPosition = screenPosition
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取选择位置失败: {ex.Message}");
            }
        }
    }

    public class TextSelectedEventArgs : EventArgs
    {
        public string SelectedText { get; set; }
        public Point ScreenPosition { get; set; }
    }
}
