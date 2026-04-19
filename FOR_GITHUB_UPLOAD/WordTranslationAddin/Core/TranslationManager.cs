using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordTranslationAddin.Models;
using WordTranslationAddin.Translation;

namespace WordTranslationAddin.Core
{
    public class TranslationManager
    {
        private readonly SettingsManager _settingsManager;
        private readonly Dictionary<TranslationEngineType, ITranslationEngine> _engines;
        private readonly OfflineDictionaryEngine _offlineEngine;

        public TranslationManager(SettingsManager settingsManager)
        {
            _settingsManager = settingsManager ?? throw new ArgumentNullException(nameof(settingsManager));
            _engines = new Dictionary<TranslationEngineType, ITranslationEngine>();
            _offlineEngine = new OfflineDictionaryEngine();

            InitializeEngines();
        }

        private void InitializeEngines()
        {
            var settings = _settingsManager.GetSettings();

            if (!string.IsNullOrEmpty(settings.GoogleApiKey))
            {
                _engines[TranslationEngineType.Google] = new GoogleTranslateEngine(settings.GoogleApiKey);
            }

            if (!string.IsNullOrEmpty(settings.BaiduAppId) && !string.IsNullOrEmpty(settings.BaiduSecretKey))
            {
                _engines[TranslationEngineType.Baidu] = new BaiduTranslateEngine(settings.BaiduAppId, settings.BaiduSecretKey);
            }

            if (!string.IsNullOrEmpty(settings.DeepLApiKey))
            {
                _engines[TranslationEngineType.DeepL] = new DeepLTranslateEngine(settings.DeepLApiKey);
            }
        }

        public async Task<TranslationResult> TranslateAsync(string text)
        {
            var settings = _settingsManager.GetSettings();
            var engineType = settings.PreferredEngine;

            // 首先尝试离线翻译（如果是单词）
            if (IsWord(text) && settings.EnableOfflineMode)
            {
                var offlineResult = await _offlineEngine.TranslateAsync(text, settings.SourceLanguage, settings.TargetLanguage);
                if (offlineResult != null && !string.IsNullOrEmpty(offlineResult.TranslatedText))
                {
                    offlineResult.EngineType = TranslationEngineType.Offline;
                    offlineResult.IsOffline = true;
                    return offlineResult;
                }
            }

            // 使用在线翻译引擎
            if (_engines.TryGetValue(engineType, out var engine))
            {
                try
                {
                    var result = await engine.TranslateAsync(text, settings.SourceLanguage, settings.TargetLanguage);
                    result.EngineType = engineType;
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"翻译引擎 {engineType} 失败: {ex.Message}");
                    
                    // 尝试使用其他引擎
                    foreach (var backupEngine in _engines)
                    {
                        if (backupEngine.Key != engineType)
                        {
                            try
                            {
                                var result = await backupEngine.Value.TranslateAsync(text, settings.SourceLanguage, settings.TargetLanguage);
                                result.EngineType = backupEngine.Key;
                                result.IsFallback = true;
                                return result;
                            }
                            catch { }
                        }
                    }
                }
            }

            // 如果所有引擎都失败，返回错误
            return new TranslationResult
            {
                SourceText = text,
                TranslatedText = "翻译失败，请检查网络连接或API配置",
                EngineType = TranslationEngineType.None,
                IsError = true
            };
        }

        private bool IsWord(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim();
            return !text.Contains(" ") && !text.Contains("\t") && !text.Contains("\n");
        }

        public void UpdateEngineConfig(TranslationEngineType engineType, string apiKey, string secretKey = null)
        {
            switch (engineType)
            {
                case TranslationEngineType.Google:
                    _engines[engineType] = new GoogleTranslateEngine(apiKey);
                    break;
                case TranslationEngineType.Baidu:
                    _engines[engineType] = new BaiduTranslateEngine(apiKey, secretKey);
                    break;
                case TranslationEngineType.DeepL:
                    _engines[engineType] = new DeepLTranslateEngine(apiKey);
                    break;
            }
        }
    }
}
