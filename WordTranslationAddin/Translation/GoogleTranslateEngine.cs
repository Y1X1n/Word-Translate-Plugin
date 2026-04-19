using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Translation
{
    public class GoogleTranslateEngine : ITranslationEngine
    {
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private const string BaseUrl = "https://translation.googleapis.com/language/translate/v2";

        public string EngineName => "Google Translate";
        public bool IsAvailable => !string.IsNullOrEmpty(_apiKey);

        public GoogleTranslateEngine(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
        }

        public async Task<TranslationResult> TranslateAsync(string text, string sourceLanguage, string targetLanguage)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentException("文本不能为空", nameof(text));

            try
            {
                string encodedText = HttpUtility.UrlEncode(text);
                string url = $"{BaseUrl}?key={_apiKey}&q={encodedText}&source={sourceLanguage}&target={targetLanguage}&format=text";

                var response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string jsonResponse = await response.Content.ReadAsStringAsync();
                var result = ParseResponse(jsonResponse, text);
                
                return result;
            }
            catch (HttpRequestException ex)
            {
                throw new Exception($"Google翻译API请求失败: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Google翻译失败: {ex.Message}");
            }
        }

        private TranslationResult ParseResponse(string jsonResponse, string sourceText)
        {
            try
            {
                var json = JObject.Parse(jsonResponse);
                var data = json["data"];
                var translations = data?["translations"] as JArray;

                if (translations != null && translations.Count > 0)
                {
                    var translation = translations[0];
                    string translatedText = translation["translatedText"]?.ToString();
                    string detectedSourceLanguage = translation["detectedSourceLanguage"]?.ToString();

                    return new TranslationResult
                    {
                        SourceText = sourceText,
                        TranslatedText = translatedText,
                        SourceLanguage = detectedSourceLanguage ?? "auto",
                        TargetLanguage = "zh",
                        EngineType = TranslationEngineType.Google,
                        Timestamp = DateTime.Now
                    };
                }

                throw new Exception("无法解析翻译响应");
            }
            catch (Exception ex)
            {
                throw new Exception($"解析Google翻译响应失败: {ex.Message}");
            }
        }
    }
}
