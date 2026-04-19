using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Translation
{
    public class DeepLTranslateEngine : ITranslationEngine
    {
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private const string BaseUrl = "https://api-free.deepl.com/v2/translate";
        private const string ProBaseUrl = "https://api.deepl.com/v2/translate";

        public string EngineName => "DeepL";
        public bool IsAvailable => !string.IsNullOrEmpty(_apiKey);

        public DeepLTranslateEngine(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("DeepL-Auth-Key", _apiKey);
        }

        public async Task<TranslationResult> TranslateAsync(string text, string sourceLanguage, string targetLanguage)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentException("文本不能为空", nameof(text));

            try
            {
                var requestBody = new
                {
                    text = new[] { text },
                    source_lang = sourceLanguage?.ToUpper() ?? "EN",
                    target_lang = ConvertLanguageCode(targetLanguage),
                    preserve_formatting = true
                };

                string jsonContent = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                string url = _apiKey.EndsWith(":fx") ? BaseUrl : ProBaseUrl;
                
                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                string jsonResponse = await response.Content.ReadAsStringAsync();
                var result = ParseResponse(jsonResponse, text);
                
                return result;
            }
            catch (HttpRequestException ex)
            {
                throw new Exception($"DeepL API请求失败: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"DeepL翻译失败: {ex.Message}");
            }
        }

        private string ConvertLanguageCode(string code)
        {
            switch (code.ToLower())
            {
                case "zh": return "ZH";
                case "zh-cn": return "ZH";
                case "zh-tw": return "ZH";
                case "en": return "EN";
                case "ja": return "JA";
                case "ko": return "KO";
                case "fr": return "FR";
                case "de": return "DE";
                case "es": return "ES";
                case "ru": return "RU";
                case "it": return "IT";
                case "pt": return "PT";
                case "nl": return "NL";
                case "pl": return "PL";
                default: return code.ToUpper();
            }
        }

        private TranslationResult ParseResponse(string jsonResponse, string sourceText)
        {
            try
            {
                var json = JObject.Parse(jsonResponse);
                var translations = json["translations"] as JArray;

                if (translations != null && translations.Count > 0)
                {
                    var translation = translations[0];
                    string translatedText = translation["text"]?.ToString();
                    string detectedSourceLanguage = translation["detected_source_language"]?.ToString();

                    return new TranslationResult
                    {
                        SourceText = sourceText,
                        TranslatedText = translatedText,
                        SourceLanguage = detectedSourceLanguage?.ToLower() ?? "auto",
                        TargetLanguage = "zh",
                        EngineType = TranslationEngineType.DeepL,
                        Timestamp = DateTime.Now
                    };
                }

                throw new Exception("无法解析DeepL翻译响应");
            }
            catch (Exception ex)
            {
                throw new Exception($"解析DeepL翻译响应失败: {ex.Message}");
            }
        }
    }
}
