using System.Threading.Tasks;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Translation
{
    public interface ITranslationEngine
    {
        Task<TranslationResult> TranslateAsync(string text, string sourceLanguage, string targetLanguage);
        bool IsAvailable { get; }
        string EngineName { get; }
    }

    public enum TranslationEngineType
    {
        None,
        Google,
        Baidu,
        DeepL,
        Offline
    }
}
