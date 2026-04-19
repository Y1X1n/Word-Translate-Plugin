using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordTranslationAddin.Helpers
{
    public static class TextHelper
    {
        /// <summary>
        /// 判断文本是否为单词（而非句子）
        /// </summary>
        public static bool IsWord(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim();
            
            // 如果包含空格、制表符或换行符，则不是单词
            if (text.Contains(' ') || text.Contains('\t') || text.Contains('\n'))
                return false;

            // 如果长度超过50个字符，可能不是单词
            if (text.Length > 50)
                return false;

            // 检查是否主要由字母组成
            int letterCount = text.Count(char.IsLetter);
            return letterCount > 0 && (double)letterCount / text.Length > 0.5;
        }

        /// <summary>
        /// 判断文本是否为句子
        /// </summary>
        public static bool IsSentence(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim();

            // 包含空格且长度适中
            if (!text.Contains(' '))
                return false;

            // 检查是否包含句子结束符
            char[] sentenceEndings = { '.', '!', '?', '。', '！', '？' };
            bool hasEnding = text.IndexOfAny(sentenceEndings) >= 0;

            // 或者包含多个单词
            int wordCount = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;

            return hasEnding || wordCount >= 3;
        }

        /// <summary>
        /// 清理文本，移除多余空白字符
        /// </summary>
        public static string CleanText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // 替换多个连续空白为单个空格
            text = Regex.Replace(text, @"\s+", " ");
            
            // 移除首尾空白
            text = text.Trim();

            return text;
        }

        /// <summary>
        /// 截断文本到指定长度
        /// </summary>
        public static string TruncateText(string text, int maxLength, string suffix = "...")
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;

            return text.Substring(0, maxLength - suffix.Length) + suffix;
        }

        /// <summary>
        /// 检测文本语言（简单实现）
        /// </summary>
        public static string DetectLanguage(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "unknown";

            // 检测中文字符
            if (Regex.IsMatch(text, @"[\u4e00-\u9fa5]"))
                return "zh";

            // 检测日文字符
            if (Regex.IsMatch(text, @"[\u3040-\u309f\u30a0-\u30ff]"))
                return "ja";

            // 检测韩文字符
            if (Regex.IsMatch(text, @"[\uac00-\ud7af]"))
                return "ko";

            // 检测俄文字符
            if (Regex.IsMatch(text, @"[\u0400-\u04ff]"))
                return "ru";

            // 默认假设为英文
            if (Regex.IsMatch(text, @"[a-zA-Z]"))
                return "en";

            return "unknown";
        }

        /// <summary>
        /// 计算文本的单词数
        /// </summary>
        public static int CountWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            return text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        /// <summary>
        /// 获取文本的字符数（不包括空白）
        /// </summary>
        public static int CountCharacters(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;

            return text.Count(c => !char.IsWhiteSpace(c));
        }

        /// <summary>
        /// 检查文本是否只包含标点符号
        /// </summary>
        public static bool IsOnlyPunctuation(string text)
        {
            if (string.IsNullOrEmpty(text))
                return true;

            return text.All(c => char.IsPunctuation(c) || char.IsWhiteSpace(c));
        }

        /// <summary>
        /// 移除文本中的HTML标签
        /// </summary>
        public static string RemoveHtmlTags(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return Regex.Replace(text, @"<[^>]+>", string.Empty);
        }
    }
}
