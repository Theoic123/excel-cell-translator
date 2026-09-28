using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace ExcelCellTranslator
{
    internal static class ProviderTests
    {
        private sealed class CapturedRequest
        {
            internal TranslationProvider Provider;
            internal string Uri;
            internal string Method;
            internal string Body;
            internal IDictionary<string, string> Headers;
        }

        internal static int Run(StringBuilder log)
        {
            int passed = 0;
            var previousHandler = TranslationService.RequestHandlerForTests;
            var previousSettings = TranslationService.SettingsForTests;
            try
            {
                TranslationService.RequestHandlerForTests = null;
                TranslationService.SettingsForTests = null;

                Check(ref passed, log,
                    TranslationService.NormalizeIdentifiers("Region_Score_Limit") == "Region Score Limit",
                    "Normalize underscore identifiers");
                Check(ref passed, log,
                    TranslationService.NormalizeIdentifiers("RegionScoreLimit") == "Region Score Limit",
                    "Normalize camel-case identifiers");
                Check(ref passed, log,
                    TranslationService.NormalizeIdentifiers("myHTTPServerName") == "my HTTP Server Name",
                    "Normalize acronym boundaries in identifiers");
                Check(ref passed, log,
                    TranslationService.NormalizeIdentifiers("profit -10 USD on 2026-09-28") == "profit -10 USD on 2026-09-28",
                    "Identifier normalization preserves negative values and dates");

                Check(ref passed, log,
                    TranslationService.ParseGoogleResponse("{\"data\":{\"translations\":[{\"translatedText\":\"译文\"}]}}") == "译文",
                    "Parse Google Cloud response");
                Check(ref passed, log,
                    TranslationService.ParseDeepLResponse("{\"translations\":[{\"text\":\"translation\"}]}") == "translation",
                    "Parse DeepL response");
                ExpectInvalid(ref passed, log,
                    delegate { TranslationService.ParseGoogleResponse("{\"error\":{\"code\":429,\"message\":\"quota\"}}"); },
                    "Reject Google provider error response");
                ExpectInvalid(ref passed, log,
                    delegate { TranslationService.ParseDeepLResponse("{\"translations\":[]}"); },
                    "Reject empty DeepL response");

                TestCredentialPersistence(ref passed, log);
                TestProviderPayloads(ref passed, log);
                TestMyMemoryPayloadAndFallback(ref passed, log);
                TestFallbackAndAuthError(ref passed, log);
                TestInFlightCancellation(ref passed, log);
            }
            finally
            {
                TranslationService.RequestHandlerForTests = previousHandler;
                TranslationService.SettingsForTests = previousSettings;
            }
            return passed;
        }

        private static void TestCredentialPersistence(ref int passed, StringBuilder log)
        {
            string directory = Path.Combine(Path.GetTempPath(), "ExcelCellTranslator-provider-tests-" + Guid.NewGuid().ToString("N"));
            string path = Path.Combine(directory, "provider-settings.dat");
            const string googleKey = "Google-test-secret-47c38d";
            const string deepLKey = "DeepL-test-secret-82f20a";
            try
            {
                Directory.CreateDirectory(directory);
                var settings = new ProviderSettings {
                    PrimaryProvider = TranslationProvider.GoogleCloud,
                    FallbackProvider = TranslationProvider.DeepL,
                    GoogleApiKey = googleKey,
                    DeepLApiKey = deepLKey,
                    DeepLUsePro = true
                };
                settings.Save(path);

                string persisted = File.ReadAllText(path, Encoding.UTF8);
                Check(ref passed, log,
                    persisted.IndexOf(googleKey, StringComparison.Ordinal) < 0 && persisted.IndexOf(deepLKey, StringComparison.Ordinal) < 0,
                    "Provider settings do not persist API keys as plaintext");

                ProviderSettings loaded = ProviderSettings.Load(path);
                Check(ref passed, log,
                    loaded.PrimaryProvider == TranslationProvider.GoogleCloud && loaded.FallbackProvider == TranslationProvider.DeepL,
                    "Provider settings round-trip provider selection");
                Check(ref passed, log,
                    loaded.GoogleApiKey == googleKey && loaded.DeepLApiKey == deepLKey && loaded.DeepLUsePro,
                    "Provider settings round-trip encrypted credentials and DeepL mode");
                Check(ref passed, log,
                    loaded.IsProviderConfigured(TranslationProvider.GoogleCloud) && loaded.IsProviderConfigured(TranslationProvider.DeepL) &&
                    loaded.IsProviderConfigured(TranslationProvider.MyMemory),
                    "Provider configuration recognizes both keys and keyless MyMemory");
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        private static void TestProviderPayloads(ref int passed, StringBuilder log)
        {
            var requests = new List<CapturedRequest>();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.GoogleCloud,
                GoogleApiKey = "Google-request-secret"
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                requests.Add(new CapturedRequest { Provider = provider, Uri = uri, Method = method, Body = body, Headers = headers });
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                    "{\"data\":{\"translations\":[{\"translatedText\":\"google result\"}]}}"));
            };
            TranslationResult google = TranslationService.TranslateAsync("provider payload sample", CancellationToken.None).GetAwaiter().GetResult();
            Check(ref passed, log, google.Text == "google result", "Use Google Cloud response translation");
            Check(ref passed, log,
                requests.Count == 1 && requests[0].Provider == TranslationProvider.GoogleCloud && requests[0].Method == "POST",
                "Route primary request to Google Cloud with POST");
            Check(ref passed, log,
                requests.Count == 1 && RequestContains(requests[0], "provider payload sample") && RequestContains(requests[0], "Google-request-secret"),
                "Google request includes source text and configured API key");
            Check(ref passed, log,
                requests.Count == 1 && requests[0].Uri == "https://translation.googleapis.com/language/translate/v2" &&
                HasHeader(requests[0], "X-Goog-Api-Key", "Google-request-secret") &&
                JsonHasValue(requests[0].Body, "q", "provider payload sample") &&
                JsonHasValue(requests[0].Body, "source", "en") &&
                JsonHasValue(requests[0].Body, "target", "zh-CN") &&
                JsonHasValue(requests[0].Body, "format", "text"),
                "Google request uses the expected endpoint, auth header, and language payload");

            requests.Clear();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.DeepL,
                DeepLApiKey = "DeepL-request-secret",
                DeepLUsePro = false
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                requests.Add(new CapturedRequest { Provider = provider, Uri = uri, Method = method, Body = body, Headers = headers });
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                    "{\"translations\":[{\"text\":\"deepl result\"}]}"));
            };
            TranslationResult deepL = TranslationService.TranslateAsync("provider payload sample for translation", CancellationToken.None).GetAwaiter().GetResult();
            Check(ref passed, log, deepL.Text == "deepl result", "Use DeepL response translation");
            Check(ref passed, log,
                requests.Count == 1 && requests[0].Provider == TranslationProvider.DeepL && requests[0].Method == "POST",
                "Route primary request to DeepL with POST");
            Check(ref passed, log,
                requests.Count == 1 && RequestContains(requests[0], "provider payload sample for translation") && RequestContains(requests[0], "DeepL-request-secret"),
                "DeepL request includes source text and configured API key");
            Check(ref passed, log,
                requests.Count == 1 && requests[0].Uri.IndexOf("api-free.deepl.com", StringComparison.OrdinalIgnoreCase) >= 0,
                "DeepL Free mode uses the Free API endpoint");
            Check(ref passed, log,
                requests.Count == 1 && requests[0].Uri == "https://api-free.deepl.com/v2/translate" &&
                HasHeader(requests[0], "Authorization", "DeepL-Auth-Key DeepL-request-secret") &&
                JsonArrayHasValue(requests[0].Body, "text", "provider payload sample for translation") &&
                JsonHasValue(requests[0].Body, "target_lang", "ZH"),
                "DeepL request uses the Free endpoint, auth header, and target payload");

            requests.Clear();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.DeepL,
                DeepLApiKey = "DeepL-pro-request-secret",
                DeepLUsePro = true
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                requests.Add(new CapturedRequest { Provider = provider, Uri = uri, Method = method, Body = body, Headers = headers });
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                    "{\"translations\":[{\"text\":\"pro result\"}]}"));
            };
            TranslationService.TranslateAsync("DeepL Pro endpoint sample", CancellationToken.None).GetAwaiter().GetResult();
            Check(ref passed, log,
                requests.Count == 1 && requests[0].Provider == TranslationProvider.DeepL &&
                requests[0].Uri == "https://api.deepl.com/v2/translate" &&
                HasHeader(requests[0], "Authorization", "DeepL-Auth-Key DeepL-pro-request-secret"),
                "DeepL Pro mode uses the Pro endpoint and configured key");
        }

        private static void TestMyMemoryPayloadAndFallback(ref int passed, StringBuilder log)
        {
            var requests = new List<CapturedRequest>();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.MyMemory,
                FallbackProvider = TranslationProvider.GoogleCloud,
                GoogleApiKey = "Google-MyMemory-fallback-secret"
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                requests.Add(new CapturedRequest { Provider = provider, Uri = uri, Method = method, Body = body, Headers = headers });
                if (provider == TranslationProvider.MyMemory)
                    return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                        "{\"responseStatus\":200,\"quotaFinished\":true}"));
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                    "{\"data\":{\"translations\":[{\"translatedText\":\"MyMemory fallback result\"}]}}"));
            };
            TranslationResult result = TranslationService.TranslateAsync("mymemory get payload sample", CancellationToken.None).GetAwaiter().GetResult();
            Check(ref passed, log,
                result.Text == "MyMemory fallback result" && requests.Count == 2 &&
                requests[0].Provider == TranslationProvider.MyMemory && requests[0].Method == "GET" &&
                requests[0].Uri.StartsWith("https://api.mymemory.translated.net/get?", StringComparison.Ordinal) &&
                RequestContains(requests[0], "mymemory get payload sample") &&
                requests[1].Provider == TranslationProvider.GoogleCloud,
                "MyMemory GET request encodes source text and falls back after quota exhaustion");
        }

        private static void TestFallbackAndAuthError(ref int passed, StringBuilder log)
        {
            var calls = new List<TranslationProvider>();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.GoogleCloud,
                FallbackProvider = TranslationProvider.DeepL,
                GoogleApiKey = "Google-fallback-secret",
                DeepLApiKey = "DeepL-fallback-secret"
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                calls.Add(provider);
                if (provider == TranslationProvider.GoogleCloud)
                    return Task.FromResult(new TranslationHttpResponse((HttpStatusCode)429,
                        "{\"error\":{\"code\":429,\"message\":\"quota exceeded\"}}"));
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                    "{\"translations\":[{\"text\":\"fallback result\"}]}"));
            };
            TranslationResult fallback = TranslationService.TranslateAsync("fallback route sample", CancellationToken.None).GetAwaiter().GetResult();
            Check(ref passed, log,
                fallback.Text == "fallback result" && calls.Count == 2 && calls[0] == TranslationProvider.GoogleCloud && calls[1] == TranslationProvider.DeepL,
                "Quota error routes through configured fallback provider");

            calls.Clear();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.DeepL,
                FallbackProvider = TranslationProvider.GoogleCloud,
                DeepLApiKey = "DeepL-456-secret",
                GoogleApiKey = "Google-456-secret"
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                calls.Add(provider);
                if (provider == TranslationProvider.DeepL)
                    return Task.FromResult(new TranslationHttpResponse((HttpStatusCode)456, "{\"message\":\"quota exceeded\"}"));
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.OK,
                    "{\"data\":{\"translations\":[{\"translatedText\":\"Google fallback result\"}]}}"));
            };
            TranslationResult deepLFallback = TranslationService.TranslateAsync("DeepL 456 quota fallback sample", CancellationToken.None).GetAwaiter().GetResult();
            Check(ref passed, log,
                deepLFallback.Text == "Google fallback result" && calls.Count == 2 &&
                calls[0] == TranslationProvider.DeepL && calls[1] == TranslationProvider.GoogleCloud,
                "DeepL 456 quota response routes through configured Google fallback");

            calls.Clear();
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.GoogleCloud,
                FallbackProvider = TranslationProvider.DeepL,
                GoogleApiKey = "Google-fallback-secret",
                DeepLApiKey = "DeepL-fallback-secret"
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                calls.Add(provider);
                return Task.FromResult(new TranslationHttpResponse(HttpStatusCode.Forbidden,
                    "{\"error\":{\"code\":403,\"message\":\"invalid API key\"}}"));
            };
            ExpectInvalid(ref passed, log,
                delegate { TranslationService.TranslateAsync("auth error does not fall back", CancellationToken.None).GetAwaiter().GetResult(); },
                "Reject invalid provider credentials");
            Check(ref passed, log,
                calls.Count == 1 && calls[0] == TranslationProvider.GoogleCloud,
                "Authentication errors do not route to fallback");
        }

        private static void TestInFlightCancellation(ref int passed, StringBuilder log)
        {
            var entered = new ManualResetEvent(false);
            TranslationService.SettingsForTests = new ProviderSettings {
                PrimaryProvider = TranslationProvider.GoogleCloud,
                GoogleApiKey = "Google-cancel-secret"
            };
            TranslationService.RequestHandlerForTests = delegate(TranslationProvider provider, string uri, string method,
                string body, IDictionary<string, string> headers, CancellationToken token)
            {
                entered.Set();
                var completion = new TaskCompletionSource<TranslationHttpResponse>();
                token.Register(delegate { completion.TrySetCanceled(); });
                return completion.Task;
            };

            using (var cancellation = new CancellationTokenSource())
            {
                Task<TranslationResult> task = TranslationService.TranslateAsync("cancel an in-flight request", cancellation.Token);
                try
                {
                    Check(ref passed, log, entered.WaitOne(TimeSpan.FromSeconds(5)), "Cancellation test reaches provider transport");
                    cancellation.Cancel();
                    bool cancelled = false;
                    try { task.GetAwaiter().GetResult(); }
                    catch (OperationCanceledException) { cancelled = true; }
                    Check(ref passed, log, cancelled, "Cancellation propagates from in-flight provider request");
                }
                finally
                {
                    cancellation.Cancel();
                    entered.Close();
                }
            }
        }

        private static bool RequestContains(CapturedRequest request, string value)
        {
            if (request == null) return false;
            if (ContainsDecoded(request.Uri, value) || ContainsDecoded(request.Body, value)) return true;
            if (request.Headers != null)
                foreach (KeyValuePair<string, string> header in request.Headers)
                    if (ContainsDecoded(header.Key, value) || ContainsDecoded(header.Value, value)) return true;
            return false;
        }

        private static bool HasHeader(CapturedRequest request, string key, string value)
        {
            if (request == null || request.Headers == null) return false;
            foreach (KeyValuePair<string, string> header in request.Headers)
                if (string.Equals(header.Key, key, StringComparison.OrdinalIgnoreCase) && header.Value == value) return true;
            return false;
        }

        private static bool JsonHasValue(string json, string key, string expected)
        {
            try
            {
                var values = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>;
                object actual;
                return values != null && values.TryGetValue(key, out actual) && string.Equals(actual as string, expected, StringComparison.Ordinal);
            }
            catch (ArgumentException) { return false; }
        }

        private static bool JsonArrayHasValue(string json, string key, string expected)
        {
            try
            {
                var values = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>;
                object raw;
                if (values == null || !values.TryGetValue(key, out raw)) return false;
                object[] items = raw as object[];
                if (items == null) return false;
                foreach (object item in items)
                    if (string.Equals(item as string, expected, StringComparison.Ordinal)) return true;
                return false;
            }
            catch (ArgumentException) { return false; }
        }

        private static bool ContainsDecoded(string source, string value)
        {
            if (string.IsNullOrEmpty(source)) return false;
            string decoded = Uri.UnescapeDataString(source.Replace('+', ' '));
            return decoded.IndexOf(value, StringComparison.Ordinal) >= 0;
        }

        private static void ExpectInvalid(ref int passed, StringBuilder log, Action action, string name)
        {
            try { action(); }
            catch (InvalidOperationException) { Check(ref passed, log, true, name); return; }
            throw new Exception("FAIL: accepted " + name);
        }

        private static void Check(ref int passed, StringBuilder log, bool value, string name)
        {
            if (!value) throw new Exception("FAIL: " + name);
            passed++;
            log.AppendLine("PASS: " + name);
        }
    }
}
