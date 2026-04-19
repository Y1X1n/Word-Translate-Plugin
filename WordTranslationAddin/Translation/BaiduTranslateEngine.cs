using System;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Newtonsoft.Json.Linq;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Translation
{
    public class BaiduTranslateEngine : ITranslationEngine
    {
        private readonly string _appId;
        private readonly string _secretKey;
        private readonly HttpClient _httpClient;
        private const string BaseUrl = "https://fanyi-api.baidu.com/api/trans/vip/translate";

        public string EngineName => "百度翻译";
        public bool IsAvailable => !string.IsNullOrEmpty(_appId) && !string.IsNullOrEmpty(_secretKey);

        public BaiduTranslateEngine(string appId, string secretKey)
        {
            _appId = appId ?? throw new ArgumentNullException(nameof(appId));
            _secretKey = secretKey ?? throw new ArgumentNullException(nameof(secretKey));
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
        }

        public async Task<TranslationResult> TranslateAsync(string text, string sourceLanguage, string targetLanguage)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentException("文本不能为空", nameof(text));

            try
            {
                string salt = DateTime.Now.Millisecond.ToString();
                string sign = GenerateSign(text, salt);
                
                string from = ConvertLanguageCode(sourceLanguage);
                string to = ConvertLanguageCode(targetLanguage);

                string url = $"{BaseUrl}?q={HttpUtility.UrlEncode(text)}&from={from}&to={to}&appid={_appId}&salt={salt}&sign={sign}";

                var response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string jsonResponse = await response.Content.ReadAsStringAsync();
                var result = ParseResponse(jsonResponse, text);
                
                return result;
            }
            catch (HttpRequestException ex)
            {
                throw new Exception($"百度翻译API请求失败: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"百度翻译失败: {ex.Message}");
            }
        }

        private string GenerateSign(string text, string salt)
        {
            string str = _appId + text + salt + _secretKey;
            using (MD5 md5 = MD5.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(str);
                byte[] hash = md5.ComputeHash(bytes);
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("x2"));
                }
                return sb.ToString();
            }
        }

        private string ConvertLanguageCode(string code)
        {
            switch (code.ToLower())
            {
                case "en": return "en";
                case "zh": return "zh";
                case "zh-cn": return "zh";
                case "zh-tw": return "cht";
                case "ja": return "jp";
                case "ko": return "kor";
                case "fr": return "fra";
                case "de": return "de";
                case "es": return "spa";
                case "ru": return "ru";
                case "auto": return "auto";
                default: return code;
            }
        }

        private TranslationResult ParseResponse(string jsonResponse, string sourceText)
        {
            try
            {
                var json = JObject.Parse(jsonResponse);
                
                // 检查错误
                var errorCode = json["error_code"]?.ToString();
                if (!string.IsNullOrEmpty(errorCode))
                {
                    var errorMsg = json["error_msg"]?.ToString();
                    throw new Exception($"百度翻译API错误: {errorCode} - {errorMsg}");
                }

                var transResult = json["trans_result"] as JArray;
                if (transResult != null && transResult.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (var item in transResult)
                    {
                        sb.Append(item["dst"]?.ToString());
                    }

                    return new TranslationResult
                    {
                        SourceText = sourceText,
                        TranslatedText = sb.ToString(),
                        SourceLanguage = json["from"]?.ToString() ?? "auto",
                        TargetLanguage = json["to"]?.ToString() ?? "zh",
                        EngineType = TranslationEngineType.Baidu,
                        Timestamp = DateTime.Now
                    };
                }

                throw new Exception("无法解析百度翻译响应");
            }
            catch (Exception ex)
            {
                throw new Exception($"解析百度翻译响应失败: {ex.Message}");
            }
        }
    }
}
