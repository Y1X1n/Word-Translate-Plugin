using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Core
{
    public class HistoryManager
    {
        private readonly string _historyFilePath;
        private List<TranslationHistory> _history;
        private readonly object _lockObject = new object();
        private readonly int _maxHistoryItems;

        public HistoryManager()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string pluginFolder = Path.Combine(appDataPath, "WordTranslationAddin");
            
            if (!Directory.Exists(pluginFolder))
            {
                Directory.CreateDirectory(pluginFolder);
            }

            _historyFilePath = Path.Combine(pluginFolder, "history.json");
            _maxHistoryItems = 100;
            LoadHistory();
        }

        private void LoadHistory()
        {
            try
            {
                if (File.Exists(_historyFilePath))
                {
                    string json = File.ReadAllText(_historyFilePath);
                    _history = JsonConvert.DeserializeObject<List<TranslationHistory>>(json) ?? new List<TranslationHistory>();
                }
                else
                {
                    _history = new List<TranslationHistory>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载历史记录失败: {ex.Message}");
                _history = new List<TranslationHistory>();
            }
        }

        public void SaveHistory()
        {
            lock (_lockObject)
            {
                try
                {
                    string json = JsonConvert.SerializeObject(_history, Formatting.Indented);
                    File.WriteAllText(_historyFilePath, json);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"保存历史记录失败: {ex.Message}");
                }
            }
        }

        public void AddToHistory(TranslationResult result)
        {
            if (result == null || result.IsError)
                return;

            lock (_lockObject)
            {
                var existingItem = _history.FirstOrDefault(h => 
                    h.SourceText.Equals(result.SourceText, StringComparison.OrdinalIgnoreCase));

                if (existingItem != null)
                {
                    existingItem.TranslatedText = result.TranslatedText;
                    existingItem.EngineType = result.EngineType;
                    existingItem.Timestamp = DateTime.Now;
                    existingItem.AccessCount++;
                }
                else
                {
                    var historyItem = new TranslationHistory
                    {
                        Id = Guid.NewGuid(),
                        SourceText = result.SourceText,
                        TranslatedText = result.TranslatedText,
                        SourceLanguage = result.SourceLanguage,
                        TargetLanguage = result.TargetLanguage,
                        EngineType = result.EngineType,
                        Timestamp = DateTime.Now,
                        AccessCount = 1
                    };

                    _history.Insert(0, historyItem);

                    if (_history.Count > _maxHistoryItems)
                    {
                        _history = _history.Take(_maxHistoryItems).ToList();
                    }
                }

                SaveHistory();
            }
        }

        public List<TranslationHistory> GetHistory(int count = 50)
        {
            lock (_lockObject)
            {
                return _history.Take(count).ToList();
            }
        }

        public List<TranslationHistory> SearchHistory(string keyword)
        {
            lock (_lockObject)
            {
                if (string.IsNullOrWhiteSpace(keyword))
                    return _history.ToList();

                return _history.Where(h => 
                    h.SourceText.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                    h.TranslatedText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }
        }

        public void ClearHistory()
        {
            lock (_lockObject)
            {
                _history.Clear();
                SaveHistory();
            }
        }

        public void RemoveFromHistory(Guid id)
        {
            lock (_lockObject)
            {
                var item = _history.FirstOrDefault(h => h.Id == id);
                if (item != null)
                {
                    _history.Remove(item);
                    SaveHistory();
                }
            }
        }
    }
}
