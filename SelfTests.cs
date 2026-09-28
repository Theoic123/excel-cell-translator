using System;
using System.IO;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelCellTranslator
{
    internal static class SelfTests
    {
        private static int passed;
        private static readonly StringBuilder Log = new StringBuilder();
        private static void Check(bool value, string name)
        {
            if (!value) throw new Exception("FAIL: " + name);
            passed++;
            Log.AppendLine("PASS: " + name);
        }
        private static void Reject(Action action, string name)
        {
            try { action(); } catch (InvalidOperationException) { Check(true, name); return; }
            throw new Exception("FAIL: accepted " + name);
        }
        public static int Run()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Check(!TranslationService.IsChinese("Hello world"), "English detection");
                Check(TranslationService.IsChinese("你好世界"), "Chinese detection");
                Check(TranslationService.IsChinese("Excel 单元格"), "Mixed text chooses English target");
                Check(TranslationService.IsChinese(char.ConvertFromUtf32(0x20000)), "Supplementary CJK detection");
                Check(!TranslationService.IsChinese("Hello \U0001F600"), "Emoji is not Chinese");
                passed += ProviderTests.Run(Log);
                passed += EditSelectionTests.Run(Log);
                string source = new string('中', 300) + " word boundary.\r\n" + new string('x', 510) + "\U0001F600 end";
                var parts = TranslationService.SplitText(source);
                Check(string.Concat(parts) == source, "Chunking preserves entire input");
                var strict = new UTF8Encoding(false, true);
                foreach (string part in parts) Check(strict.GetByteCount(part) <= 480, "Chunk obeys UTF-8 byte limit");
                Check(TranslationService.SplitText("").Count == 0, "Empty split terminates");
                string emojiSource = new string('x', 479) + "\U0001F600 next";
                foreach (string part in TranslationService.SplitText(emojiSource)) Check(strict.GetByteCount(part) <= 480, "Never split surrogate pair");
                Check(TranslationService.ParseResponse("{\"responseStatus\":200,\"responseData\":{\"translatedText\":\"Hello &amp; world\"},\"quotaFinished\":false}") == "Hello & world", "Parse and decode response");
                Check(TranslationService.ParseResponse("{\"responseStatus\":\"200\",\"responseData\":{\"translatedText\":\"你好\"}}") == "你好", "String service status");
                Reject(delegate { TranslationService.ParseResponse("{\"responseStatus\":200,\"quotaFinished\":true}"); }, "Quota exhausted");
                Reject(delegate { TranslationService.ParseResponse("{\"responseStatus\":429}"); }, "Service failure");
                Reject(delegate { TranslationService.ParseResponse("{\"responseStatus\":200,\"responseData\":{\"translatedText\":\"\"}}"); }, "Empty translation");
                Reject(delegate { TranslationService.ParseResponse("null"); }, "Null payload");
                Reject(delegate { TranslationService.ParseResponse("not-json"); }, "Malformed payload");
                Reject(delegate { TranslationService.TranslateAsync("  ", CancellationToken.None).GetAwaiter().GetResult(); }, "Empty source rejected before network");
                Reject(delegate { TranslationService.TranslateAsync(new string('a', 2001), CancellationToken.None).GetAwaiter().GetResult(); }, "Oversize rejected before network");
                using (var cancelled = new CancellationTokenSource())
                {
                    cancelled.Cancel();
                    bool rejected = false;
                    try { TranslationService.TranslateAsync("Hello", cancelled.Token).GetAwaiter().GetResult(); }
                    catch (OperationCanceledException) { rejected = true; }
                    Check(rejected, "Cancelled request never starts network");
                }
                using (var main = new MainForm())
                {
                    main.ShowInTaskbar = false;
                    main.Show();
                    Application.DoEvents();
                    foreach (Control control in main.Controls)
                        Check(control.Bottom <= main.ClientSize.Height && control.Right <= main.ClientSize.Width,
                            "Main window control is within client area: " + control.GetType().Name);
                    using (var bitmap = new Bitmap(main.Width, main.Height))
                    {
                        main.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                        bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "preview.png"));
                    }
                }
                using (var providerSettingsForm = new ProviderSettingsForm())
                {
                    providerSettingsForm.ShowInTaskbar = false;
                    providerSettingsForm.Show();
                    Application.DoEvents();
                    foreach (Control control in providerSettingsForm.Controls)
                        Check(control.Bottom <= providerSettingsForm.ClientSize.Height && control.Right <= providerSettingsForm.ClientSize.Width,
                            "Provider settings control is within client area: " + control.GetType().Name);
                    using (var bitmap = new Bitmap(providerSettingsForm.Width, providerSettingsForm.Height))
                    {
                        providerSettingsForm.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                        bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "service-settings-preview.png"));
                    }
                }
                using (var popup = new TranslationPopupForm("English → 中文", "选中一个文字单元格，右键即可翻译。切换选区后，译文会自动消失。", false))
                {
                    popup.Show();
                    Application.DoEvents();
                    using (var bitmap = new Bitmap(popup.Width, popup.Height))
                    {
                        popup.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                        bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "popup-preview.png"));
                    }
                }
                using (var popup = new TranslationPopupForm("中文 → English", new string('文', 2000), false))
                {
                    bool scrollable = false;
                    foreach (Control control in popup.Controls)
                    {
                        var panel = control as Panel;
                        if (panel != null && panel.AutoScroll && panel.AutoScrollMinSize.Height > panel.Height) scrollable = true;
                    }
                    Check(scrollable, "Long translation has a scrollable body");
                }
                Log.AppendLine("ALL PASSED: " + passed);
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "self-test.txt"), Log.ToString(), Encoding.UTF8);
                return 0;
            }
            catch (Exception error)
            {
                Log.AppendLine(error.ToString());
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "self-test.txt"), Log.ToString(), Encoding.UTF8);
                return 1;
            }
        }
        public static int RunExcel() { return ExcelSmokeTests.Run(); }
    }
}
