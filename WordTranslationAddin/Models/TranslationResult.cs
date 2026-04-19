using System;
using System.Collections.Generic;
using WordTranslationAddin.Translation;

namespace WordTranslationAddin.Models
{
    public class TranslationResult
    {
        public string SourceText { get; set; }
        public string TranslatedText { get; set; }
        public string SourceLanguage { get; set; }
        public string TargetLanguage { get; set; }
        public TranslationEngineType EngineType { get; set; }
        public bool IsOffline { get; set; }
        public bool IsFallback { get; set; }
        public bool IsError { get; set; }
        
        // 单词详细信息（仅单词翻译时有）
        public string Phonetic { get; set; }
        public string PartOfSpeech { get; set; }
        public string Definition { get; set; }
        public List<string> Examples { get; set; }
        public List<string> Synonyms { get; set; }
        
        public DateTime Timestamp { get; set; }
        public long ResponseTimeMs { get; set; }

        public bool IsWord => !string.IsNullOrEmpty(Phonetic) || 
                              (!string.IsNullOrEmpty(SourceText) && !SourceText.Contains(" "));

        public TranslationResult()
        {
            Examples = new List<string>();
            Synonyms = new List<string>();
        }
    }
}
