using System;
using System.IO;
using Newtonsoft.Json;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Core
{
    public class SettingsManager
    {
        private readonly string _settingsFilePath;
        private AppSettings _settings;
        private readonly object _lockObject = new object();

        public SettingsManager()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string pluginFolder = Path.Combine(appDataPath, "WordTranslationAddin");
            
            if (!Directory.Exists(pluginFolder))
            {
                Directory.CreateDirectory(pluginFolder);
            }

            _settingsFilePath = Path.Combine(pluginFolder, "settings.json");
            LoadSettings();
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    _settings = JsonConvert.DeserializeObject<AppSettings>(json);
                }
                else
                {
                    _settings = GetDefaultSettings();
                    SaveSettings();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载设置失败: {ex.Message}");
                _settings = GetDefaultSettings();
            }
        }

        public void SaveSettings()
        {
            lock (_lockObject)
            {
                try
                {
                    string json = JsonConvert.SerializeObject(_settings, Formatting.Indented);
                    File.WriteAllText(_settingsFilePath, json);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"保存设置失败: {ex.Message}");
                }
            }
        }

        public AppSettings GetSettings()
        {
            lock (_lockObject)
            {
                return _settings;
            }
        }

        public void UpdateSettings(Action<AppSettings> updateAction)
        {
            lock (_lockObject)
            {
                updateAction(_settings);
                SaveSettings();
            }
        }

        private AppSettings GetDefaultSettings()
        {
            return new AppSettings
            {
                PreferredEngine = TranslationEngineType.Google,
                SourceLanguage = "en",
                TargetLanguage = "zh",
                EnableOfflineMode = true,
                PopupWidth = 400,
                PopupHeight = 300,
                PopupPosition = PopupPosition.Auto,
                EnableAnimation = true,
                AutoCopyTranslation = false,
                ShowPhonetic = true,
                ShowExamples = true,
                MaxHistoryItems = 100,
                Theme = AppTheme.System,
                ShortcutKey = "Ctrl+Shift+T",
                EnableAutoTranslate = true,
                MinTextLength = 1,
                MaxTextLength = 500
            };
        }
    }
}
