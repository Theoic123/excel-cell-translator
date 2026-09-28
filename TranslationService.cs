using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace ExcelCellTranslator
{
    internal sealed class TranslationResult
    {
        public string Text { get; private set; }
        public string Direction { get; private set; }
        public string Provider { get; private set; }

        public TranslationResult(string text, string direction) : this(text, direction, string.Empty) { }
        public TranslationResult(string text, string direction, string provider)
        {
            Text = text;
            Direction = direction;
            Provider = provider ?? string.Empty;
        }
    }

    internal sealed class TranslationHttpResponse
    {
        public HttpStatusCode StatusCode { get; private set; }
        public string Body { get; private set; }

        internal TranslationHttpResponse(HttpStatusCode statusCode, string body)
        {
            StatusCode = statusCode;
            Body = body ?? string.Empty;
        }
    }

    internal sealed class TranslationProviderException : InvalidOperationException
    {
        internal bool CanFallback { get; private set; }

        internal TranslationProviderException(string message, bool canFallback)
            : base(message)
        {
            CanFallback = canFallback;
        }
    }

    internal static class TranslationService
    {
        internal const int MaximumCharacters = 2000;
        private const int MaximumResponseCharacters = 262144;
        private const int RequestTimeoutSeconds = 15;
        private const string MyMemoryEndpoint = "https://api.mymemory.translated.net/get";
        private const string GoogleEndpoint = "https://translation.googleapis.com/language/translate/v2";
        private const string DeepLFreeEndpoint = "https://api-free.deepl.com/v2/translate";
        private const string DeepLProEndpoint = "https://api.deepl.com/v2/translate";

        private static readonly Dictionary<string, TranslationResult> Cache = new Dictionary<string, TranslationResult>();
        private static readonly Queue<string> CacheOrder = new Queue<string>();
        private static readonly object CacheLock = new object();

        // Test overrides are internal so production callers continue to use persisted settings and HTTP.
        internal static ProviderSettings SettingsForTests { get; set; }
        internal static Func<TranslationProvider, string, string, string, IDictionary<string, string>, CancellationToken, Task<TranslationHttpResponse>> RequestHandlerForTests { get; set; }

        internal static bool IsChinese(string text)
        {
            if (text == null) return false;
            for (int i = 0; i < text.Length; i++)
            {
                int code = char.ConvertToUtf32(text, i);
                if ((code >= 0x3400 && code <= 0x9fff) || (code >= 0xf900 && code <= 0xfaff) ||
                    (code >= 0x20000 && code <= 0x323af)) return true;
                if (char.IsHighSurrogate(text[i])) i++;
            }
            return false;
        }

        // Turn common snake_case, kebab-case, PascalCase and camelCase identifiers into words.
        // The caller supplies the returned value to the provider; the source cell itself is untouched.
        internal static string NormalizeIdentifiers(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            string normalized = Regex.Replace(text, @"([a-z0-9])([A-Z])", "$1 $2");
            normalized = Regex.Replace(normalized, @"([A-Z])([A-Z][a-z])", "$1 $2");
            normalized = Regex.Replace(normalized, @"_+", " ");
            normalized = Regex.Replace(normalized, @"(?<=[A-Za-z])-(?=[A-Za-z])", " ");
            return normalized;
        }

        // Keep each chunk below the provider's 500 UTF-8 byte limit, without splitting surrogate pairs.
        internal static List<string> SplitText(string text)
        {
            var parts = new List<string>();
            int start = 0;
            while (start < text.Length)
            {
                int end = start, bytes = 0, boundary = -1;
                while (end < text.Length)
                {
                    int count = char.IsHighSurrogate(text[end]) && end + 1 < text.Length && char.IsLowSurrogate(text[end + 1]) ? 2 : 1;
                    int size = Encoding.UTF8.GetByteCount(text.Substring(end, count));
                    if (bytes + size > 480) break;
                    bytes += size;
                    end += count;
                    char last = text[end - 1];
                    if (char.IsWhiteSpace(last) || ".!?;。！？；".IndexOf(last) >= 0) boundary = end;
                }
                if (end < text.Length && boundary > start) end = boundary;
                parts.Add(text.Substring(start, end - start));
                start = end;
            }
            return parts;
        }

        public static async Task<TranslationResult> TranslateAsync(string text, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("这个单元格没有可翻译的文字。");
            if (text.Length > MaximumCharacters) throw new InvalidOperationException("文字超过 2,000 字符，请先缩短单元格内容。");

            ProviderSettings settings = SettingsForTests ?? ProviderSettings.Load();
            if (settings == null) settings = new ProviderSettings();
            bool containsNormalizedIdentifier = !string.Equals(text, NormalizeIdentifiers(text), StringComparison.Ordinal);
            string sourceText = NormalizeIdentifiers(text);
            if (string.IsNullOrWhiteSpace(sourceText) || !ContainsLetter(sourceText))
                throw new InvalidOperationException("单元格中没有可翻译的文字，请选择包含词语的内容后重试。");
            bool chinese = IsChinese(sourceText);
            string pair = chinese ? "zh-CN|en" : "en|zh-CN";
            string direction = chinese ? "中文 → English" : "English → 中文";
            string key = BuildCacheKey(pair, text, settings);

            lock (CacheLock)
            {
                TranslationResult cached;
                if (Cache.TryGetValue(key, out cached)) return cached;
            }

            using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                deadline.CancelAfter(TimeSpan.FromSeconds(45));
                var translated = new StringBuilder();
                var usedProviders = new List<TranslationProvider>();
                try
                {
                    foreach (string part in SplitText(sourceText))
                    {
                        deadline.Token.ThrowIfCancellationRequested();
                        if (string.IsNullOrWhiteSpace(part)) { translated.Append(part); continue; }
                        ProviderTranslation result = await FetchWithFallbackAsync(part.Trim(), pair, settings, deadline.Token).ConfigureAwait(false);
                        if (!usedProviders.Contains(result.Provider)) usedProviders.Add(result.Provider);
                        if (translated.Length > 0 && !char.IsWhiteSpace(translated[translated.Length - 1])) translated.Append(' ');
                        translated.Append(result.Text);

                        // Retain source paragraph and word separators between chunks.
                        int trailing = part.Length;
                        while (trailing > 0 && char.IsWhiteSpace(part[trailing - 1])) trailing--;
                        if (trailing < part.Length) translated.Append(part.Substring(trailing));
                    }
                }
                catch (OperationCanceledException)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    throw new InvalidOperationException("翻译请求超时，请检查网络后重试。");
                }

                cancellationToken.ThrowIfCancellationRequested();
                string answerText = translated.ToString().Trim();
                if (string.IsNullOrWhiteSpace(answerText))
                    throw new InvalidOperationException("翻译服务没有返回可用的译文，请更换服务或稍后重试。");
                if (containsNormalizedIdentifier &&
                    (string.Equals(answerText, sourceText.Trim(), StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(answerText, text.Trim(), StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("翻译服务原样返回了标识符“" + sourceText.Trim() + "”。请在设置中更换翻译服务，或提供更多上下文后重试。");

                var answer = new TranslationResult(answerText, direction, GetProviderNames(usedProviders, settings.DeepLUsePro));
                lock (CacheLock)
                {
                    if (!Cache.ContainsKey(key))
                    {
                        while (CacheOrder.Count >= 64) Cache.Remove(CacheOrder.Dequeue());
                        CacheOrder.Enqueue(key);
                        Cache.Add(key, answer);
                    }
                }
                return answer;
            }
        }

        private static async Task<ProviderTranslation> FetchWithFallbackAsync(string text, string pair, ProviderSettings settings, CancellationToken token)
        {
            TranslationProvider primary = settings.PrimaryProvider;
            if (!settings.IsProviderConfigured(primary))
                throw MissingKeyError(primary);

            ProviderTranslation primaryResult = default(ProviderTranslation);
            TranslationProviderException primaryError = null;
            try { primaryResult = await FetchFromProviderAsync(text, pair, primary, settings, token).ConfigureAwait(false); }
            catch (TranslationProviderException error) { primaryError = error; }
            if (primaryError == null) return primaryResult;

            TranslationProvider? fallback = settings.FallbackProvider;
            if (!primaryError.CanFallback || !fallback.HasValue || fallback.Value == primary) throw primaryError;
            if (!settings.IsProviderConfigured(fallback.Value))
                throw new InvalidOperationException(primaryError.Message + "已配置的备用服务“" + ProviderSettings.DisplayName(fallback.Value, settings.DeepLUsePro) + "”缺少 API 密钥。", primaryError);

            try { return await FetchFromProviderAsync(text, pair, fallback.Value, settings, token).ConfigureAwait(false); }
            catch (TranslationProviderException fallbackError)
            {
                throw new InvalidOperationException(primaryError.Message + "已尝试备用服务“" + ProviderSettings.DisplayName(fallback.Value, settings.DeepLUsePro) + "”，但仍无法完成翻译：" + fallbackError.Message, fallbackError);
            }
        }

        private static async Task<ProviderTranslation> FetchFromProviderAsync(string text, string pair, TranslationProvider provider, ProviderSettings settings, CancellationToken token)
        {
            string url;
            string method;
            string body = string.Empty;
            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            switch (provider)
            {
                case TranslationProvider.MyMemory:
                    url = MyMemoryEndpoint + "?q=" + Uri.EscapeDataString(text) + "&langpair=" + Uri.EscapeDataString(pair);
                    method = "GET";
                    break;
                case TranslationProvider.GoogleCloud:
                    url = GoogleEndpoint;
                    method = "POST";
                    headers["X-Goog-Api-Key"] = settings.GoogleApiKey;
                    headers["Content-Type"] = "application/json; charset=utf-8";
                    body = SerializeGoogleRequest(text, pair);
                    break;
                case TranslationProvider.DeepL:
                    url = settings.DeepLUsePro ? DeepLProEndpoint : DeepLFreeEndpoint;
                    method = "POST";
                    headers["Authorization"] = "DeepL-Auth-Key " + settings.DeepLApiKey;
                    headers["Content-Type"] = "application/json; charset=utf-8";
                    body = SerializeDeepLRequest(text, pair);
                    break;
                default:
                    throw new InvalidOperationException("所选的翻译服务无效。");
            }

            TranslationHttpResponse response = await SendRequestAsync(provider, url, method, body, headers, token).ConfigureAwait(false);
            int status = (int)response.StatusCode;
            if (status != 200)
            {
                bool eligible = IsTransientHttpStatus(status) ||
                    (provider == TranslationProvider.DeepL && status == 456) ||
                    (provider == TranslationProvider.GoogleCloud && GoogleBodyIsQuotaError(response.Body));
                throw new TranslationProviderException(GetHttpErrorMessage(provider, status, eligible), eligible);
            }

            string translated;
            switch (provider)
            {
                case TranslationProvider.MyMemory: translated = ParseResponse(response.Body); break;
                case TranslationProvider.GoogleCloud: translated = WebUtility.HtmlDecode(ParseGoogleResponse(response.Body)); break;
                case TranslationProvider.DeepL: translated = ParseDeepLResponse(response.Body); break;
                default: throw new InvalidOperationException("所选的翻译服务无效。");
            }
            return new ProviderTranslation(translated, provider);
        }

        private static string SerializeGoogleRequest(string text, string pair)
        {
            string[] languages = pair.Split('|');
            var payload = new Dictionary<string, object>();
            payload["q"] = text;
            payload["source"] = languages[0];
            payload["target"] = languages[1];
            payload["format"] = "text";
            return new JavaScriptSerializer().Serialize(payload);
        }

        private static string SerializeDeepLRequest(string text, string pair)
        {
            bool chinese = pair.StartsWith("zh-", StringComparison.OrdinalIgnoreCase);
            var payload = new Dictionary<string, object>();
            payload["text"] = new string[] { text };
            payload["target_lang"] = chinese ? "EN-US" : "ZH";
            return new JavaScriptSerializer().Serialize(payload);
        }

        private static async Task<TranslationHttpResponse> SendRequestAsync(
            TranslationProvider provider,
            string url,
            string method,
            string body,
            IDictionary<string, string> headers,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            using (var requestDeadline = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                requestDeadline.CancelAfter(TimeSpan.FromSeconds(RequestTimeoutSeconds));
                try
                {
                    Func<TranslationProvider, string, string, string, IDictionary<string, string>, CancellationToken, Task<TranslationHttpResponse>> handler = RequestHandlerForTests;
                    if (handler != null)
                    {
                        TranslationHttpResponse testResponse = await handler(provider, url, method, body, headers, requestDeadline.Token).ConfigureAwait(false);
                        token.ThrowIfCancellationRequested();
                        requestDeadline.Token.ThrowIfCancellationRequested();
                        if (testResponse == null) throw new TranslationProviderException("翻译服务没有返回响应。", false);
                        return testResponse;
                    }

                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                    var request = (HttpWebRequest)WebRequest.Create(url);
                    request.Method = method;
                    request.UserAgent = "ExcelCellTranslator/1.0";
                    request.Accept = "application/json";
                    request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                    request.AllowAutoRedirect = false;
                    foreach (KeyValuePair<string, string> header in headers)
                    {
                        if (string.Equals(header.Key, "Content-Type", StringComparison.OrdinalIgnoreCase)) request.ContentType = header.Value;
                        else request.Headers[header.Key] = header.Value;
                    }

                    using (requestDeadline.Token.Register(request.Abort))
                    {
                        if (!string.IsNullOrEmpty(body))
                        {
                            byte[] bytes = Encoding.UTF8.GetBytes(body);
                            request.ContentLength = bytes.Length;
                            using (Stream requestStream = await request.GetRequestStreamAsync().ConfigureAwait(false))
                            {
                                requestDeadline.Token.ThrowIfCancellationRequested();
                                await requestStream.WriteAsync(bytes, 0, bytes.Length, requestDeadline.Token).ConfigureAwait(false);
                            }
                        }

                        HttpWebResponse response = null;
                        try
                        {
                            response = (HttpWebResponse)await request.GetResponseAsync().ConfigureAwait(false);
                        }
                        catch (WebException error)
                        {
                            response = error.Response as HttpWebResponse;
                            if (response == null) throw;
                        }

                        using (response)
                        using (Stream stream = response.GetResponseStream())
                        {
                            string responseBody = stream == null ? string.Empty : await ReadResponseAsync(stream, requestDeadline.Token).ConfigureAwait(false);
                            return new TranslationHttpResponse(response.StatusCode, responseBody);
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    token.ThrowIfCancellationRequested();
                    if (requestDeadline.IsCancellationRequested)
                        throw new TranslationProviderException("翻译服务请求超时。", true);
                    throw;
                }
                catch (WebException error)
                {
                    token.ThrowIfCancellationRequested();
                    if (requestDeadline.IsCancellationRequested)
                        throw new TranslationProviderException("翻译服务请求超时。", true);
                    if (error.Response != null) error.Response.Dispose();
                    throw new TranslationProviderException("无法连接翻译服务，请检查网络后重试。", true);
                }
            }
        }

        private static async Task<string> ReadResponseAsync(Stream stream, CancellationToken token)
        {
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                var body = new StringBuilder();
                var buffer = new char[2048];
                int count;
                while ((count = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
                {
                    token.ThrowIfCancellationRequested();
                    body.Append(buffer, 0, count);
                    if (body.Length > MaximumResponseCharacters) throw new TranslationProviderException("翻译服务返回的数据过大。", false);
                }
                return body.ToString();
            }
        }

        private static string GetHttpErrorMessage(TranslationProvider provider, int status, bool canFallback)
        {
            if (provider == TranslationProvider.DeepL && status == 456)
                return "DeepL API 配额或用量限制已达到。";
            if (provider == TranslationProvider.GoogleCloud && (status == 429 || status == 403 && canFallback))
                return "Google Cloud Translation 配额或请求频率限制已达到。";
            if (status == 429)
            {
                string service = provider == TranslationProvider.DeepL ? "DeepL API" : "MyMemory";
                return service + "请求频率受限，请稍后重试。";
            }
            if (status >= 500 || status == 408)
                return "翻译服务暂时不可用（HTTP " + status.ToString(System.Globalization.CultureInfo.InvariantCulture) + "）。";
            return "翻译服务拒绝了请求（HTTP " + status.ToString(System.Globalization.CultureInfo.InvariantCulture) + "），请检查密钥和服务配置。";
        }

        private static bool IsTransientHttpStatus(int status)
        {
            return status == 408 || status == 429 || (status >= 500 && status <= 599);
        }

        internal static string ParseResponse(string json)
        {
            Dictionary<string, object> root = ParseObject(json);
            object value;
            if (root.TryGetValue("quotaFinished", out value) && object.Equals(value, true))
                throw new TranslationProviderException("今日 MyMemory 免费翻译额度已用完，请明天再试或更换服务。", true);

            if (!root.TryGetValue("responseStatus", out value))
                throw new InvalidOperationException("MyMemory 返回的数据缺少服务状态。");
            string status = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
            if (status != "200")
            {
                object detailsValue;
                string details = root.TryGetValue("responseDetails", out detailsValue) ? Convert.ToString(detailsValue, System.Globalization.CultureInfo.InvariantCulture) : string.Empty;
                bool eligible = status == "408" || status == "429" || status == "503" || status.StartsWith("5", StringComparison.Ordinal);
                if (!eligible && !string.IsNullOrEmpty(details))
                    eligible = details.IndexOf("quota", StringComparison.OrdinalIgnoreCase) >= 0 || details.IndexOf("limit", StringComparison.OrdinalIgnoreCase) >= 0;
                throw new TranslationProviderException(eligible ? "MyMemory 免费翻译额度或服务暂时不可用。" : "MyMemory 未能完成翻译请求。", eligible);
            }

            Dictionary<string, object> data = root.TryGetValue("responseData", out value) ? value as Dictionary<string, object> : null;
            string result = data != null && data.TryGetValue("translatedText", out value) ? value as string : null;
            if (string.IsNullOrWhiteSpace(result)) throw new InvalidOperationException("MyMemory 没有返回译文，请稍后重试。");
            return WebUtility.HtmlDecode(result);
        }

        internal static string ParseGoogleResponse(string json)
        {
            Dictionary<string, object> root = ParseObject(json);
            object value;
            if (root.TryGetValue("error", out value))
            {
                Dictionary<string, object> error = value as Dictionary<string, object>;
                string status = GetString(error, "status");
                int code = GetInt(error, "code");
                bool quota = IsGoogleQuotaStatus(status) || IsGoogleQuotaReason(GetGoogleErrorReason(error)) || code == 429 || code == 403 && IsGoogleQuotaReason(GetGoogleErrorMessage(error));
                bool eligible = quota || IsTransientHttpStatus(code);
                throw new TranslationProviderException(eligible ? "Google Cloud Translation 配额或服务暂时不可用。" : "Google Cloud Translation 请求失败，请检查 API 密钥和 Cloud Translation 配置。", eligible);
            }

            Dictionary<string, object> data = root.TryGetValue("data", out value) ? value as Dictionary<string, object> : null;
            object[] translations = data != null && data.TryGetValue("translations", out value) ? value as object[] : null;
            Dictionary<string, object> first = translations != null && translations.Length > 0 ? translations[0] as Dictionary<string, object> : null;
            string result = GetString(first, "translatedText");
            if (string.IsNullOrWhiteSpace(result)) throw new InvalidOperationException("Google Cloud Translation 没有返回译文。");
            return result;
        }

        internal static string ParseDeepLResponse(string json)
        {
            Dictionary<string, object> root = ParseObject(json);
            object value;
            if (root.TryGetValue("translations", out value))
            {
                object[] translations = value as object[];
                Dictionary<string, object> first = translations != null && translations.Length > 0 ? translations[0] as Dictionary<string, object> : null;
                string result = GetString(first, "text");
                if (!string.IsNullOrWhiteSpace(result)) return result;
            }

            string code = GetString(root, "code");
            string message = GetString(root, "message");
            bool quota = string.Equals(code, "456", StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrEmpty(message) && message.IndexOf("quota", StringComparison.OrdinalIgnoreCase) >= 0);
            if (quota) throw new TranslationProviderException("DeepL API 配额或用量限制已达到。", true);
            throw new InvalidOperationException("DeepL 没有返回译文，请检查 API 密钥和服务配置。");
        }

        private static Dictionary<string, object> ParseObject(string json)
        {
            Dictionary<string, object> root;
            try { root = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>; }
            catch (ArgumentException) { throw new InvalidOperationException("翻译服务返回的数据无法识别。"); }
            catch (InvalidOperationException) { throw new InvalidOperationException("翻译服务返回的数据无法识别。"); }
            if (root == null) throw new InvalidOperationException("翻译服务返回的数据无法识别。");
            return root;
        }

        private static bool GoogleBodyIsQuotaError(string json)
        {
            try
            {
                Dictionary<string, object> root = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>;
                object value;
                Dictionary<string, object> error = root != null && root.TryGetValue("error", out value) ? value as Dictionary<string, object> : null;
                if (error == null) return false;
                return IsGoogleQuotaStatus(GetString(error, "status")) ||
                    IsGoogleQuotaReason(GetGoogleErrorReason(error)) ||
                    GetInt(error, "code") == 403 && IsGoogleQuotaReason(GetGoogleErrorMessage(error));
            }
            catch (Exception error)
            {
                if (error is OutOfMemoryException || error is StackOverflowException) throw;
                return false;
            }
        }

        private static string GetGoogleErrorReason(Dictionary<string, object> error)
        {
            if (error == null) return string.Empty;
            object value;
            object[] details = error.TryGetValue("errors", out value) ? value as object[] : null;
            if (details == null) return string.Empty;
            var reasons = new StringBuilder();
            foreach (object item in details)
            {
                Dictionary<string, object> detail = item as Dictionary<string, object>;
                string reason = GetString(detail, "reason");
                if (reason.Length > 0) reasons.Append(' ').Append(reason);
            }
            return reasons.ToString();
        }

        private static string GetGoogleErrorMessage(Dictionary<string, object> error)
        {
            return GetString(error, "message");
        }

        private static bool IsGoogleQuotaStatus(string status)
        {
            return string.Equals(status, "RESOURCE_EXHAUSTED", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(status, "QUOTA_EXCEEDED", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsGoogleQuotaReason(string reason)
        {
            if (string.IsNullOrEmpty(reason)) return false;
            string lower = reason.ToLowerInvariant();
            return lower.Contains("quota") || lower.Contains("ratelimit") || lower.Contains("rate_limit") || lower.Contains("limitexceeded");
        }

        private static string GetString(Dictionary<string, object> source, string name)
        {
            object value;
            return source != null && source.TryGetValue(name, out value) ? Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) : string.Empty;
        }

        private static int GetInt(Dictionary<string, object> source, string name)
        {
            object value;
            int result;
            return source != null && source.TryGetValue(name, out value) && int.TryParse(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture), out result) ? result : 0;
        }

        private static InvalidOperationException MissingKeyError(TranslationProvider provider)
        {
            switch (provider)
            {
                case TranslationProvider.GoogleCloud: return new InvalidOperationException("请先在设置中填写 Google Cloud Translation API 密钥。");
                case TranslationProvider.DeepL: return new InvalidOperationException("请先在设置中填写 DeepL API 密钥。");
                default: return new InvalidOperationException("所选翻译服务尚未正确配置。");
            }
        }

        private static string GetProviderNames(List<TranslationProvider> providers, bool deepLUsePro)
        {
            var names = new List<string>();
            foreach (TranslationProvider provider in providers)
            {
                string name = ProviderSettings.DisplayName(provider, deepLUsePro);
                if (!names.Contains(name)) names.Add(name);
            }
            return string.Join(" / ", names.ToArray());
        }

        private static bool ContainsLetter(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsLetter(text, i)) return true;
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) i++;
            }
            return false;
        }

        private static string BuildCacheKey(string pair, string text, ProviderSettings settings)
        {
            string credentials = (settings.GoogleApiKey ?? string.Empty) + "\0" + (settings.DeepLApiKey ?? string.Empty);
            byte[] bytes = Encoding.UTF8.GetBytes(credentials);
            byte[] hash = null;
            try
            {
                using (SHA256 sha = SHA256.Create()) hash = sha.ComputeHash(bytes);
                string fallback = settings.FallbackProvider.HasValue ? settings.FallbackProvider.Value.ToString() : "-";
                return pair + "\n" + settings.PrimaryProvider.ToString() + "\n" + fallback + "\n" + settings.DeepLUsePro.ToString() + "\n" + Convert.ToBase64String(hash) + "\n" + text;
            }
            finally
            {
                Array.Clear(bytes, 0, bytes.Length);
                if (hash != null) Array.Clear(hash, 0, hash.Length);
            }
        }

        private struct ProviderTranslation
        {
            internal readonly string Text;
            internal readonly TranslationProvider Provider;
            internal ProviderTranslation(string text, TranslationProvider provider) { Text = text; Provider = provider; }
        }
    }
}
