using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace ExcelCellTranslator
{
    internal static class EditSelectionTests
    {
        internal static int Run(StringBuilder log)
        {
            int passed = 0;
            TestNativeInputLayout(ref passed, log);
            TestAccessiblePartialSelection(ref passed, log);
            TestClipboardSnapshots(ref passed, log);
            TestClipboardUpdateClaim(ref passed, log);
            TestNativeMenuEntrypoint(ref passed, log);
            return passed;
        }

        private static void TestNativeInputLayout(ref int passed, StringBuilder log)
        {
            Type input = typeof(EditSelectionBridge).GetNestedType("Input", BindingFlags.NonPublic);
            Type keyboard = typeof(EditSelectionBridge).GetNestedType("KeyboardInput", BindingFlags.NonPublic);
            Check(ref passed, log, Marshal.SizeOf(input) == (IntPtr.Size == 8 ? 40 : 28),
                "Copy INPUT matches the Win32 ABI size, including the complete union");
            Check(ref passed, log, Marshal.OffsetOf(input, "Data").ToInt32() == (IntPtr.Size == 8 ? 8 : 4)
                && Marshal.OffsetOf(keyboard, "ExtraInfo").ToInt32() == (IntPtr.Size == 8 ? 16 : 12),
                "Copy INPUT preserves native alignment and injected-key marker offset");
        }

        private static void TestAccessiblePartialSelection(ref int passed, StringBuilder log)
        {
            using (var form = new Form())
            using (var edit = new RichTextBox())
            using (var process = System.Diagnostics.Process.GetCurrentProcess())
            {
                form.ShowInTaskbar = false;
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(20, 20);
                form.ClientSize = new Size(420, 100);
                edit.Dock = DockStyle.Fill;
                string content = "prefix: chosen partial text :suffix";
                string expected = "chosen partial text";
                edit.Text = content;
                form.Controls.Add(edit);
                form.Show();
                edit.Focus();
                edit.Select(content.IndexOf(expected, StringComparison.Ordinal), expected.Length);
                Application.DoEvents();

                AutomationElement element = AutomationElement.FromHandle(edit.Handle);
                bool selectionCompleted;
                string selected = ReadSelectionWithPump(element, (uint)process.Id, out selectionCompleted);
                Check(ref passed, log, selectionCompleted && selected == expected,
                    "UI Automation returns only the RichTextBox partial selection");

                edit.Select(0, 0);
                Application.DoEvents();
                bool emptyCompleted;
                string emptySelection = ReadSelectionWithPump(element, (uint)process.Id, out emptyCompleted);
                Check(ref passed, log, emptyCompleted && emptySelection == null,
                    "UI Automation does not substitute the whole RichTextBox value for an empty selection");
            }
        }

        private static void TestClipboardSnapshots(ref int passed, StringBuilder log)
        {
            TestStringSnapshot(ref passed, log);
            TestArraySnapshot(ref passed, log);
            TestByteSnapshot(ref passed, log);
            TestMemoryStreamSnapshot(ref passed, log);
            TestBitmapSnapshot(ref passed, log);
            TestEmptySnapshot(ref passed, log);
            var unsafeData = new DataObject();
            unsafeData.SetData(DataFormats.UnicodeText, false, "existing text is preserved");
            unsafeData.SetData("ExcelCellTranslator.Test.Unsafe", false, new object());
            ExpectInvalid(ref passed, log, delegate { SnapshotUnsafeData(unsafeData); },
                "Unsafe clipboard format is rejected before copy");
            Check(ref passed, log,
                (unsafeData.GetData(DataFormats.UnicodeText, false) as string) == "existing text is preserved",
                "Rejecting unsafe clipboard data leaves the source object unchanged");
        }

        private static void TestStringSnapshot(ref int passed, StringBuilder log)
        {
            const string value = "clipboard text remains available";
            var original = new DataObject();
            original.SetData(DataFormats.UnicodeText, false, value);
            using (EditSelectionBridge.ClipboardSnapshot snapshot = EditSelectionBridge.SnapshotClipboardData(original))
            {
                string copied = snapshot.DataObject.GetData(DataFormats.UnicodeText, false) as string;
                string source = original.GetData(DataFormats.UnicodeText, false) as string;
                Check(ref passed, log,
                    !snapshot.IsEmpty && copied == value && source == value && !object.ReferenceEquals(original, snapshot.DataObject),
                    "Clipboard snapshot preserves text in an independent data object");
            }
        }

        private static void TestArraySnapshot(ref int passed, StringBuilder log)
        {
            string[] files = { "source.xlsx", "notes.txt" };
            var original = new DataObject();
            original.SetData(DataFormats.FileDrop, false, files);
            using (EditSelectionBridge.ClipboardSnapshot snapshot = EditSelectionBridge.SnapshotClipboardData(original))
            {
                files[0] = "changed.xlsx";
                string[] copied = snapshot.DataObject.GetData(DataFormats.FileDrop, false) as string[];
                Check(ref passed, log,
                    copied != null && !object.ReferenceEquals(files, copied) && copied.Length == 2 && copied[0] == "source.xlsx",
                    "Clipboard snapshot clones string-array data");
            }
        }

        private static void TestByteSnapshot(ref int passed, StringBuilder log)
        {
            const string format = "ExcelCellTranslator.Test.Bytes";
            byte[] bytes = { 17, 34, 51 };
            var original = new DataObject();
            original.SetData(format, false, bytes);
            using (EditSelectionBridge.ClipboardSnapshot snapshot = EditSelectionBridge.SnapshotClipboardData(original))
            {
                bytes[0] = 99;
                byte[] copied = snapshot.DataObject.GetData(format, false) as byte[];
                Check(ref passed, log,
                    copied != null && !object.ReferenceEquals(bytes, copied) && copied.Length == 3 && copied[0] == 17,
                    "Clipboard snapshot clones byte-array data");
            }
        }

        private static void TestMemoryStreamSnapshot(ref int passed, StringBuilder log)
        {
            const string format = "ExcelCellTranslator.Test.MemoryStream";
            byte[] content = { 3, 5, 7, 11 };
            var source = new MemoryStream(content);
            source.Position = 2;
            var original = new DataObject();
            original.SetData(format, false, source);
            MemoryStream copied = null;
            try
            {
                using (EditSelectionBridge.ClipboardSnapshot snapshot = EditSelectionBridge.SnapshotClipboardData(original))
                {
                    copied = snapshot.DataObject.GetData(format, false) as MemoryStream;
                    Check(ref passed, log,
                        copied != null && !object.ReferenceEquals(source, copied) && copied.ToArray().Length == 4 &&
                        copied.ToArray()[0] == 3 && source.Position == 2,
                        "Clipboard snapshot clones streams and restores source position");
                }
                Check(ref passed, log, copied != null && !copied.CanRead,
                    "Disposing a clipboard snapshot disposes its owned stream clone");
            }
            finally { source.Dispose(); }
        }

        private static void TestBitmapSnapshot(ref int passed, StringBuilder log)
        {
            var source = new Bitmap(2, 2);
            source.SetPixel(0, 0, Color.Crimson);
            var original = new DataObject();
            original.SetData(DataFormats.Bitmap, false, source);
            Bitmap copied = null;
            try
            {
                using (EditSelectionBridge.ClipboardSnapshot snapshot = EditSelectionBridge.SnapshotClipboardData(original))
                {
                    copied = snapshot.DataObject.GetData(DataFormats.Bitmap, false) as Bitmap;
                    Check(ref passed, log,
                        copied != null && !object.ReferenceEquals(source, copied) && copied.GetPixel(0, 0).ToArgb() == Color.Crimson.ToArgb(),
                        "Clipboard snapshot clones bitmap data");
                }
                bool disposed = false;
                try { copied.GetPixel(0, 0); }
                catch (ArgumentException) { disposed = true; }
                catch (ObjectDisposedException) { disposed = true; }
                Check(ref passed, log, copied != null && disposed,
                    "Disposing a clipboard snapshot disposes its owned bitmap clone");
            }
            finally { source.Dispose(); }
        }

        private static void TestEmptySnapshot(ref int passed, StringBuilder log)
        {
            using (EditSelectionBridge.ClipboardSnapshot empty = EditSelectionBridge.SnapshotClipboardData(null))
                Check(ref passed, log, empty.IsEmpty && empty.DataObject == null,
                    "Null clipboard data produces an empty snapshot");
            using (EditSelectionBridge.ClipboardSnapshot emptyData = EditSelectionBridge.SnapshotClipboardData(new DataObject()))
                Check(ref passed, log, emptyData.IsEmpty && emptyData.DataObject == null,
                    "Zero-format clipboard data produces an empty snapshot");
        }

        private static void SnapshotUnsafeData(IDataObject original)
        {
            using (EditSelectionBridge.ClipboardSnapshot snapshot = EditSelectionBridge.SnapshotClipboardData(original)) { }
        }

        private static void TestNativeMenuEntrypoint(ref int passed, StringBuilder log)
        {
            Type bridgeType = typeof(EditSelectionBridge);
            MethodInfo entry = bridgeType.GetMethod("TranslateSelectionFromNativeMenu", BindingFlags.Instance | BindingFlags.NonPublic);
            Check(ref passed, log,
                entry != null && entry.ReturnType == typeof(void) && entry.GetParameters().Length == 1 &&
                entry.GetParameters()[0].ParameterType == typeof(long),
                "Native Formula Bar command entrypoint accepts the expected input revision");

            PropertyInfo revision = bridgeType.GetProperty("InputRevision", BindingFlags.Instance | BindingFlags.NonPublic);
            Check(ref passed, log, revision != null && revision.PropertyType == typeof(long),
                "Native Formula Bar command can capture the input revision");

            Check(ref passed, log,
                bridgeType.GetNestedType("EditActionForm", BindingFlags.NonPublic) == null,
                "No replacement edit action form remains");
            Check(ref passed, log,
                bridgeType.GetMethod("ReplayNativeContextMenu", BindingFlags.Instance | BindingFlags.NonPublic) == null,
                "No synthetic native context menu replay remains");

            Check(ref passed, log,
                !EditSelectionBridge.IsNativeMenuWaitCurrent(12, 12, 4, 5, 9, 9),
                "A superseded native menu wait cannot invalidate a newer translation");
        }

        private static void TestClipboardUpdateClaim(ref int passed, StringBuilder log)
        {
            const uint beforeSequence = 41;
            const uint currentSequence = 42;
            const long beforeInputRevision = 8;
            const uint expectedOwner = 1234;
            Check(ref passed, log,
                !EditSelectionBridge.CanClaimClipboardUpdate(beforeSequence, currentSequence,
                    beforeInputRevision, beforeInputRevision + 1, expectedOwner, expectedOwner),
                "Clipboard copy is not claimed after intervening user input");
            Check(ref passed, log,
                EditSelectionBridge.CanClaimClipboardUpdate(beforeSequence, currentSequence,
                    beforeInputRevision, beforeInputRevision, expectedOwner, expectedOwner),
                "Clipboard update is claimed when only the sequence changes for the expected owner");
            Check(ref passed, log,
                !EditSelectionBridge.CanClaimClipboardUpdate(beforeSequence, beforeSequence,
                    beforeInputRevision, beforeInputRevision, expectedOwner, expectedOwner),
                "Clipboard copy is not claimed when the sequence does not change");
            Check(ref passed, log,
                !EditSelectionBridge.CanClaimClipboardUpdate(beforeSequence, currentSequence,
                    beforeInputRevision, beforeInputRevision, expectedOwner, expectedOwner + 1),
                "Clipboard update is not claimed for a different owner process");
        }

        private static string ReadSelectionWithPump(AutomationElement element, uint processId, out bool completed)
        {
            Task<string> task = Task.Factory.StartNew(delegate
            {
                return EditSelectionBridge.ReadAccessibleSelection(element, processId);
            });
            DateTime deadline = DateTime.UtcNow.AddSeconds(5);
            while (!task.IsCompleted && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(10);
            }
            Application.DoEvents();
            completed = task.IsCompleted;
            return completed ? task.GetAwaiter().GetResult() : null;
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
