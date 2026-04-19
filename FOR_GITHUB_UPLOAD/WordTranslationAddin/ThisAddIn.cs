using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordTranslationAddin.Core;
using WordTranslationAddin.UI;
using System.Windows.Threading;

namespace WordTranslationAddin
{
    public partial class ThisAddIn
    {
        private TextSelectionMonitor _selectionMonitor;
        private TranslationManager _translationManager;
        private SettingsManager _settingsManager;
        private HistoryManager _historyManager;
        private TranslationPopup _translationPopup;
        private Dispatcher _dispatcher;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                _dispatcher = Dispatcher.CurrentDispatcher;
                
                _settingsManager = new SettingsManager();
                _historyManager = new HistoryManager();
                _translationManager = new TranslationManager(_settingsManager);
                
                _selectionMonitor = new TextSelectionMonitor(Application);
                _selectionMonitor.TextSelected += OnTextSelected;
                _selectionMonitor.StartMonitoring();

                Application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插件启动失败: {ex.Message}", "翻译插件", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _selectionMonitor?.StopMonitoring();
            _translationPopup?.Close();
            _historyManager?.SaveHistory();
        }

        private void Application_WindowSelectionChange(Selection sel)
        {
            _selectionMonitor?.HandleSelectionChange(sel);
        }

        private void OnTextSelected(object sender, TextSelectedEventArgs e)
        {
            _dispatcher.Invoke(() =>
            {
                ShowTranslationPopup(e.SelectedText, e.ScreenPosition);
            });
        }

        private async void ShowTranslationPopup(string text, System.Drawing.Point position)
        {
            try
            {
                if (_translationPopup != null && _translationPopup.IsVisible)
                {
                    _translationPopup.Close();
                }

                _translationPopup = new TranslationPopup(_settingsManager, _historyManager);
                _translationPopup.SetSourceText(text);
                
                var result = await _translationManager.TranslateAsync(text);
                _translationPopup.SetTranslationResult(result);
                
                _translationPopup.SetPosition(position);
                _translationPopup.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示翻译弹窗失败: {ex.Message}");
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
