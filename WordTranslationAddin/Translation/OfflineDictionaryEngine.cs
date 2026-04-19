using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Translation
{
    public class OfflineDictionaryEngine : ITranslationEngine
    {
        private Dictionary<string, OfflineWordEntry> _dictionary;
        private readonly object _lockObject = new object();
        private bool _isInitialized;

        public string EngineName => "离线词典";
        public bool IsAvailable => _isInitialized;

        public OfflineDictionaryEngine()
        {
            InitializeDictionary();
        }

        private void InitializeDictionary()
        {
            try
            {
                lock (_lockObject)
                {
                    if (_isInitialized)
                        return;

                    // 尝试从嵌入资源加载词典
                    var assembly = Assembly.GetExecutingAssembly();
                    var resourceName = "WordTranslationAddin.Resources.OfflineDictionary.json";

                    using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                    {
                        if (stream != null)
                        {
                            using (StreamReader reader = new StreamReader(stream))
                            {
                                string json = reader.ReadToEnd();
                                var entries = JsonConvert.DeserializeObject<List<OfflineWordEntry>>(json);
                                _dictionary = entries.ToDictionary(e => e.Word.ToLower(), e => e);
                            }
                        }
                        else
                        {
                            // 使用内置基础词典
                            InitializeBasicDictionary();
                        }
                    }

                    _isInitialized = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"初始化离线词典失败: {ex.Message}");
                InitializeBasicDictionary();
                _isInitialized = true;
            }
        }

        private void InitializeBasicDictionary()
        {
            _dictionary = new Dictionary<string, OfflineWordEntry>
            {
                ["hello"] = new OfflineWordEntry
                {
                    Word = "hello",
                    Translation = "你好；您好；喂",
                    Phonetic = "/həˈloʊ/",
                    PartOfSpeech = "int.",
                    Definition = "used as a greeting or to begin a telephone conversation",
                    Examples = new List<string>
                    {
                        "Hello, how are you?",
                        "Hello, is anybody there?"
                    }
                },
                ["world"] = new OfflineWordEntry
                {
                    Word = "world",
                    Translation = "世界；地球；天下",
                    Phonetic = "/wɜːrld/",
                    PartOfSpeech = "n.",
                    Definition = "the earth, together with all of its countries and peoples",
                    Examples = new List<string>
                    {
                        "He wants to travel around the world.",
                        "The whole world is watching."
                    }
                },
                ["computer"] = new OfflineWordEntry
                {
                    Word = "computer",
                    Translation = "计算机；电脑",
                    Phonetic = "/kəmˈpjuːtər/",
                    PartOfSpeech = "n.",
                    Definition = "an electronic device for storing and processing data",
                    Examples = new List<string>
                    {
                        "I work on a computer all day.",
                        "The computer crashed again."
                    }
                },
                ["translation"] = new OfflineWordEntry
                {
                    Word = "translation",
                    Translation = "翻译；译本；转化",
                    Phonetic = "/trænsˈleɪʃn/",
                    PartOfSpeech = "n.",
                    Definition = "the process of translating words or text from one language into another",
                    Examples = new List<string>
                    {
                        "The book loses something in translation.",
                        "He specializes in translation from French into English."
                    }
                },
                ["document"] = new OfflineWordEntry
                {
                    Word = "document",
                    Translation = "文件；文档；文献",
                    Phonetic = "/ˈdɑːkjumənt/",
                    PartOfSpeech = "n.",
                    Definition = "a piece of written, printed, or electronic matter that provides information",
                    Examples = new List<string>
                    {
                        "Please read the legal documents carefully.",
                        "Save the document before closing."
                    }
                },
                ["word"] = new OfflineWordEntry
                {
                    Word = "word",
                    Translation = "单词；词；字",
                    Phonetic = "/wɜːrd/",
                    PartOfSpeech = "n.",
                    Definition = "a single distinct meaningful element of speech or writing",
                    Examples = new List<string>
                    {
                        "What does this word mean?",
                        "Can I have a word with you?"
                    }
                },
                ["language"] = new OfflineWordEntry
                {
                    Word = "language",
                    Translation = "语言；语言文字",
                    Phonetic = "/ˈlæŋɡwɪdʒ/",
                    PartOfSpeech = "n.",
                    Definition = "the method of human communication, either spoken or written",
                    Examples = new List<string>
                    {
                        "English is a global language.",
                        "She speaks three languages fluently."
                    }
                },
                ["software"] = new OfflineWordEntry
                {
                    Word = "software",
                    Translation = "软件",
                    Phonetic = "/ˈsɔːftwer/",
                    PartOfSpeech = "n.",
                    Definition = "the programs and other operating information used by a computer",
                    Examples = new List<string>
                    {
                        "You need to install the software first.",
                        "The company develops educational software."
                    }
                },
                ["application"] = new OfflineWordEntry
                {
                    Word = "application",
                    Translation = "应用程序；应用；申请",
                    Phonetic = "/ˌæplɪˈkeɪʃn/",
                    PartOfSpeech = "n.",
                    Definition = "a program or piece of software designed to fulfill a particular purpose",
                    Examples = new List<string>
                    {
                        "Download the mobile application.",
                        "This application runs on Windows."
                    }
                },
                ["plugin"] = new OfflineWordEntry
                {
                    Word = "plugin",
                    Translation = "插件；外挂程序",
                    Phonetic = "/ˈplʌɡɪn/",
                    PartOfSpeech = "n.",
                    Definition = "a software component that adds a specific feature to an existing computer program",
                    Examples = new List<string>
                    {
                        "Install the plugin to enable this feature.",
                        "This browser plugin blocks ads."
                    }
                }
            };
        }

        public Task<TranslationResult> TranslateAsync(string text, string sourceLanguage, string targetLanguage)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(text))
                    return null;

                string word = text.Trim().ToLower();
                
                lock (_lockObject)
                {
                    if (_dictionary.TryGetValue(word, out var entry))
                    {
                        return new TranslationResult
                        {
                            SourceText = text,
                            TranslatedText = entry.Translation,
                            SourceLanguage = "en",
                            TargetLanguage = "zh",
                            EngineType = TranslationEngineType.Offline,
                            IsOffline = true,
                            Phonetic = entry.Phonetic,
                            PartOfSpeech = entry.PartOfSpeech,
                            Definition = entry.Definition,
                            Examples = entry.Examples,
                            Timestamp = DateTime.Now
                        };
                    }
                }

                return null;
            });
        }
    }

    public class OfflineWordEntry
    {
        public string Word { get; set; }
        public string Translation { get; set; }
        public string Phonetic { get; set; }
        public string PartOfSpeech { get; set; }
        public string Definition { get; set; }
        public List<string> Examples { get; set; }
    }
}
