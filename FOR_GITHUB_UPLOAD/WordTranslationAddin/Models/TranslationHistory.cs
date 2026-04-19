using System;
using WordTranslationAddin.Translation;

namespace WordTranslationAddin.Models
{
    public class TranslationHistory
    {
        public Guid Id { get; set; }
        public string SourceText { get; set; }
        public string TranslatedText { get; set; }
        public string SourceLanguage { get; set; }
        public string TargetLanguage { get; set; }
        public TranslationEngineType EngineType { get; set; }
        public DateTime Timestamp { get; set; }
        public int AccessCount { get; set; }

        public TranslationHistory()
        {
            Id = Guid.NewGuid();
            Timestamp = DateTime.Now;
            AccessCount = 1;
        }
    }
}
