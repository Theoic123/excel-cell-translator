using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace ExcelCellTranslator
{
    internal enum TranslationProvider
    {
        MyMemory = 0,
        GoogleCloud = 1,
        DeepL = 2
    }

    internal sealed class ProviderSettings
    {
        private const string SettingsFolderName = "ExcelCellTranslator";
        private const string SettingsFileName = "provider-settings.json";
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ExcelCellTranslator.ProviderSettings.v1");

        internal TranslationProvider PrimaryProvider { get; set; }
        internal TranslationProvider? FallbackProvider { get; set; }
        internal string GoogleApiKey { get; set; }
        internal string DeepLApiKey { get; set; }
        internal bool DeepLUsePro { get; set; }

        internal ProviderSettings()
        {
            PrimaryProvider = TranslationProvider.MyMemory;
            GoogleApiKey = string.Empty;
            DeepLApiKey = string.Empty;
        }

        internal static ProviderSettings Load()
        {
            string directory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Load(Path.Combine(directory, SettingsFolderName, SettingsFileName));
        }

        internal static ProviderSettings Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("设置文件路径不能为空。", "path");
            if (!File.Exists(path)) return new ProviderSettings();

            try
            {
                string json = File.ReadAllText(path, Encoding.UTF8);
                var data = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>;
                if (data == null) throw new InvalidDataException();

                var settings = new ProviderSettings();
                object value;
                string primary = data.TryGetValue("primaryProvider", out value) ? value as string : null;
                TranslationProvider parsedPrimary;
                if (!TryParseProvider(primary, out parsedPrimary)) throw new InvalidDataException();
                settings.PrimaryProvider = parsedPrimary;

                string fallback = data.TryGetValue("fallbackProvider", out value) ? value as string : null;
                if (!string.IsNullOrWhiteSpace(fallback))
                {
                    TranslationProvider parsedFallback;
                    if (!TryParseProvider(fallback, out parsedFallback)) throw new InvalidDataException();
                    settings.FallbackProvider = parsedFallback;
                }

                settings.DeepLUsePro = data.TryGetValue("deepLUsePro", out value) && object.Equals(value, true);
                settings.GoogleApiKey = Unprotect(data.TryGetValue("googleApiKey", out value) ? value as string : null);
                settings.DeepLApiKey = Unprotect(data.TryGetValue("deepLApiKey", out value) ? value as string : null);
                Validate(settings);
                return settings;
            }
            catch (Exception error)
            {
                if (error is OutOfMemoryException || error is StackOverflowException) throw;
                throw new InvalidOperationException("无法读取翻译服务设置。请在设置窗口重新保存服务配置。");
            }
        }

        internal void Save()
        {
            string directory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            Save(Path.Combine(directory, SettingsFolderName, SettingsFileName));
        }

        internal void Save(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("设置文件路径不能为空。", "path");
            Validate(this);

            string fullPath = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrEmpty(directory)) throw new InvalidOperationException("翻译服务设置路径无效。");
            Directory.CreateDirectory(directory);

            var data = new Dictionary<string, object>();
            data["version"] = 1;
            data["primaryProvider"] = PrimaryProvider.ToString();
            data["fallbackProvider"] = FallbackProvider.HasValue ? FallbackProvider.Value.ToString() : null;
            data["deepLUsePro"] = DeepLUsePro;
            data["googleApiKey"] = Protect(GoogleApiKey);
            data["deepLApiKey"] = Protect(DeepLApiKey);

            string temporaryPath = fullPath + ".tmp";
            try
            {
                string json = new JavaScriptSerializer().Serialize(data);
                File.WriteAllText(temporaryPath, json, new UTF8Encoding(false));
                if (File.Exists(fullPath)) File.Replace(temporaryPath, fullPath, null);
                else File.Move(temporaryPath, fullPath);
            }
            finally
            {
                if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            }
        }

        internal bool IsProviderConfigured(TranslationProvider provider)
        {
            switch (provider)
            {
                case TranslationProvider.MyMemory:
                    return true;
                case TranslationProvider.GoogleCloud:
                    return !string.IsNullOrWhiteSpace(GoogleApiKey);
                case TranslationProvider.DeepL:
                    return !string.IsNullOrWhiteSpace(DeepLApiKey);
                default:
                    return false;
            }
        }

        internal static string DisplayName(TranslationProvider provider, bool deepLUsePro)
        {
            switch (provider)
            {
                case TranslationProvider.MyMemory: return "MyMemory";
                case TranslationProvider.GoogleCloud: return "Google Cloud Translation";
                case TranslationProvider.DeepL: return deepLUsePro ? "DeepL Pro" : "DeepL Free";
                default: return "翻译服务";
            }
        }

        private static void Validate(ProviderSettings settings)
        {
            if (settings == null || !Enum.IsDefined(typeof(TranslationProvider), settings.PrimaryProvider))
                throw new InvalidOperationException("所选的主要翻译服务无效。");
            if (settings.FallbackProvider.HasValue &&
                (!Enum.IsDefined(typeof(TranslationProvider), settings.FallbackProvider.Value) ||
                 settings.FallbackProvider.Value == settings.PrimaryProvider))
                throw new InvalidOperationException("备用翻译服务必须有效，并且不能与主要服务相同。");
        }

        private static bool TryParseProvider(string value, out TranslationProvider provider)
        {
            if (Enum.TryParse<TranslationProvider>(value, false, out provider) &&
                Enum.IsDefined(typeof(TranslationProvider), provider)) return true;
            provider = TranslationProvider.MyMemory;
            return false;
        }

        private static string Protect(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            byte[] plain = Encoding.UTF8.GetBytes(value);
            byte[] encrypted = null;
            try
            {
                encrypted = ProtectedData.Protect(plain, Entropy, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            }
            finally
            {
                Array.Clear(plain, 0, plain.Length);
                if (encrypted != null) Array.Clear(encrypted, 0, encrypted.Length);
            }
        }

        private static string Unprotect(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            byte[] encrypted = null;
            byte[] plain = null;
            try
            {
                encrypted = Convert.FromBase64String(value);
                plain = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            finally
            {
                if (encrypted != null) Array.Clear(encrypted, 0, encrypted.Length);
                if (plain != null) Array.Clear(plain, 0, plain.Length);
            }
        }
    }
}
