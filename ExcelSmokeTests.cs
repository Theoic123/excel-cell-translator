using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelCellTranslator
{
    // Runs against a private Excel.Application instance and an unsaved workbook.
    // The injected translator is deterministic and never sends cell contents over the network.
    internal static class ExcelSmokeTests
    {
        private sealed class Request
        {
            public string Source;
            public TaskCompletionSource<TranslationResult> Completion;
        }

        private static readonly StringBuilder Log = new StringBuilder();
        private static int passed;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        private static void Check(bool value, string name)
        {
            if (!value) throw new InvalidOperationException("FAIL: " + name);
            passed++;
            Log.AppendLine("PASS: " + name);
            WriteLog();
        }

        private static void Phase(string message)
        {
            Log.AppendLine("PHASE: " + message);
            WriteLog();
        }

        private static uint ProcessIdForWindow(IntPtr window)
        {
            if (window == IntPtr.Zero) return 0;
            uint processId;
            GetWindowThreadProcessId(window, out processId);
            return processId;
        }

        private static uint ForegroundProcessId()
        {
            return ProcessIdForWindow(GetForegroundWindow());
        }

        private static void LogState(string label, Excel.Application excel, ExcelTranslator controller, List<Request> requests)
        {
            uint excelProcessId = 0;
            string ready;
            try
            {
                int hwnd = excel.Hwnd;
                excelProcessId = ProcessIdForWindow(new IntPtr(hwnd));
                ready = Convert.ToString(excel.Ready);
            }
            catch (Exception error) { ready = "error:" + error.GetType().Name + ":" + error.Message; }
            Log.AppendLine("STATE: " + label +
                " requests=" + requests.Count +
                " generation=" + controller.CurrentGeneration +
                " source=" + (controller.LastSourceText ?? "<null>") +
                " status=" + (controller.StatusText ?? "<null>") +
                " popupVisible=" + controller.IsPopupVisible +
                " excelReady=" + ready +
                " excelPid=" + excelProcessId +
                " foregroundPid=" + ForegroundProcessId());
            WriteLog();
        }

        private static void ReportNativeForeground(Excel.Application excel)
        {
            int hwnd = excel.Hwnd;
            uint excelProcessId = ProcessIdForWindow(new IntPtr(hwnd));
            uint foregroundProcessId = ForegroundProcessId();
            Phase("Native foreground observation only: Excel PID=" + excelProcessId + "; foreground PID=" + foregroundProcessId);
        }

        private static void PumpFor(int milliseconds)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            while (watch.ElapsedMilliseconds < milliseconds)
            {
                Application.DoEvents();
                Thread.Sleep(10);
            }
            Application.DoEvents();
        }

        private static void WaitUntil(Func<bool> condition, int timeoutMilliseconds, string operation)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            while (!condition())
            {
                if (watch.ElapsedMilliseconds >= timeoutMilliseconds)
                    throw new TimeoutException("Timed out waiting for " + operation + ".");
                Application.DoEvents();
                Thread.Sleep(10);
            }
            Application.DoEvents();
        }

        private static void SetValue(Excel.Worksheet sheet, string address, object value)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.get_Range(address, Type.Missing);
                range.Value2 = value;
            }
            finally { Release(range); }
        }

        private static object ReadValue(Excel.Worksheet sheet, string address)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.get_Range(address, Type.Missing);
                return range.Value2;
            }
            finally { Release(range); }
        }

        private static void Select(Excel.Worksheet sheet, string address)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.get_Range(address, Type.Missing);
                range.Select();
                Application.DoEvents();
            }
            finally { Release(range); }
        }

        private static void CheckPopupAnchor(Excel.Application excel, ExcelTranslator controller,
            Excel.Worksheet sheet, string address, string label)
        {
            Excel.Range range = null;
            Excel.Range hit = null;
            Excel.Window window = null;
            object hitObject = null;
            try
            {
                range = sheet.get_Range(address, Type.Missing);
                window = excel.ActiveWindow;
                double zoom = Convert.ToDouble(window.Zoom) / 100.0;
                double scale = zoom * ExcelTranslator.WindowDpi(new IntPtr(window.Hwnd)) / 72.0;
                int anchorX = window.PointsToScreenPixelsX(0) + (int)Math.Round((range.Left + range.Width) * scale, MidpointRounding.AwayFromZero);
                int anchorY = window.PointsToScreenPixelsY(0) + (int)Math.Round(range.Top * scale, MidpointRounding.AwayFromZero);
                Point expectedAnchor = new Point(anchorX, anchorY);
                Check(controller.PopupAnchor == expectedAnchor, label + " anchor matches Excel screen coordinates");

                Rectangle bounds = controller.PopupBounds;
                Phase(label + " anchor=" + expectedAnchor + "; popup=" + bounds + "; screen=" + Screen.FromPoint(expectedAnchor).WorkingArea);
                bool nextToCell = (bounds.Left >= anchorX && bounds.Left <= anchorX + 12) ||
                    (bounds.Right <= anchorX && bounds.Right >= anchorX - 12);
                Check(bounds != Rectangle.Empty && nextToCell && Math.Abs(bounds.Top - anchorY) <= 3,
                    label + " popup is positioned beside the selected cell");

                int centerX = window.PointsToScreenPixelsX(0) + (int)Math.Round((range.Left + range.Width / 2.0) * scale, MidpointRounding.AwayFromZero);
                int centerY = window.PointsToScreenPixelsY(0) + (int)Math.Round((range.Top + range.Height / 2.0) * scale, MidpointRounding.AwayFromZero);
                hitObject = window.RangeFromPoint(centerX, centerY);
                hit = hitObject as Excel.Range;
                if (hit != null)
                {
                    string expectedAddress = range.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    string actualAddress = hit.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    Check(string.Equals(actualAddress, expectedAddress, StringComparison.OrdinalIgnoreCase),
                        label + " screen coordinate maps back to the selected range");
                }
                else Check(false, label + " screen coordinate maps back to the selected range");
                Release(hit);
                hit = null;
                hitObject = window.RangeFromPoint(anchorX - 2, centerY);
                hit = hitObject as Excel.Range;
                Check(hit != null && hit.Column == range.Column, label + " right edge minus two pixels is inside selected cell");
                Release(hit);
                hit = null;
                hitObject = window.RangeFromPoint(anchorX + 2, centerY);
                hit = hitObject as Excel.Range;
                Check(hit != null && hit.Column == range.Column + 1, label + " right edge plus two pixels is inside next cell");
            }
            finally
            {
                Release(hit);
                if (!object.ReferenceEquals(hitObject, hit)) Release(hitObject);
                Release(window);
                Release(range);
            }
        }

        private static Office.CommandBarControl FindButton(Excel.Application excel, string menuTag, string barName = "Cell")
        {
            Office.CommandBars bars = null;
            Office.CommandBar cellBar = null;
            Office.CommandBarControls controls = null;
            try
            {
                bars = excel.CommandBars;
                cellBar = bars[barName];
                if (cellBar == null) return null;
                controls = cellBar.Controls;
                for (int index = 1; index <= controls.Count; index++)
                {
                    Office.CommandBarControl item = null;
                    try
                    {
                        item = controls[index];
                        if (string.Equals(item.Tag as string, menuTag, StringComparison.Ordinal))
                        {
                            Office.CommandBarControl found = item;
                            item = null;
                            return found;
                        }
                    }
                    finally { Release(item); }
                }
                return null;
            }
            finally
            {
                Release(controls);
                Release(cellBar);
                Release(bars);
            }
        }

        private static void ExecuteButton(Excel.Application excel, ExcelTranslator controller,
            string menuTag, List<Request> requests, string label)
        {
            Office.CommandBarControl button = null;
            try
            {
                ReportNativeForeground(excel);
                LogState("before " + label, excel, controller, requests);
                button = FindButton(excel, menuTag);
                if (button == null) throw new InvalidOperationException("The tagged Excel cell context-menu button was not installed.");
                button.Execute();
                LogState("after " + label, excel, controller, requests);
            }
            finally { Release(button); }
        }

        private static string NativeEditMenuItems(Excel.Application excel, string ignoredTag)
        {
            Office.CommandBars bars = null;
            Office.CommandBar bar = null;
            Office.CommandBarControls controls = null;
            var items = new List<string>();
            try
            {
                bars = excel.CommandBars;
                bar = bars["Formula Bar"];
                controls = bar.Controls;
                for (int i = 1; i <= controls.Count; i++)
                {
                    Office.CommandBarControl item = null;
                    try
                    {
                        item = controls[i];
                        if (!string.Equals(item.Tag, ignoredTag, StringComparison.Ordinal))
                            items.Add(item.Id + "|" + item.Caption + "|" + item.BeginGroup);
                    }
                    finally { Release(item); }
                }
                return string.Join("\n", items.ToArray());
            }
            finally { Release(controls); Release(bar); Release(bars); }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.ReleaseComObject(value); }
                catch { }
            }
        }

        private static void WriteLog()
        {
            try
            {
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel-test.txt"),
                    Log.ToString(), new UTF8Encoding(false));
            }
            catch (Exception error)
            {
                Log.AppendLine("Could not write excel-test.txt: " + error);
            }
        }

        public static int Run()
        {
            Log.Length = 0;
            passed = 0;
            WriteLog();
            Excel.Application excel = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            Excel.Range mergedRange = null;
            ExcelTranslator controller = null;
            string menuTag = null;
            Request delayedRequest = null;
            bool delayNextRequest = false;
            bool simulatedExcelForeground = true;
            int result = 1;
            var requests = new List<Request>();

            try
            {
                Check(Thread.CurrentThread.GetApartmentState() == ApartmentState.STA, "Excel test runs on an STA thread");
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());

                Func<string, CancellationToken, Task<TranslationResult>> fakeTranslator = delegate(string source, CancellationToken token)
                {
                    var request = new Request { Source = source };
                    requests.Add(request);
                    if (delayNextRequest)
                    {
                        delayNextRequest = false;
                        request.Completion = new TaskCompletionSource<TranslationResult>();
                        delayedRequest = request;
                        return request.Completion.Task;
                    }
                    return Task.FromResult(new TranslationResult("test translation: " + source, "test"));
                };

                Phase("before Excel.Application constructor");
                excel = new Excel.Application();
                Phase("after Excel.Application constructor");
                excel.DisplayAlerts = false;
                excel.Visible = true;
                excel.ScreenUpdating = true;
                excel.WindowState = Excel.XlWindowState.xlMaximized;
                Excel.Window activeExcelWindow = null;
                try
                {
                    activeExcelWindow = excel.ActiveWindow;
                    if (activeExcelWindow != null)
                    {
                        activeExcelWindow.Activate();
                        activeExcelWindow.Zoom = 100;
                    }
                }
                finally { Release(activeExcelWindow); }
                workbooks = excel.Workbooks;
                Phase("before Workbooks.Add");
                workbook = workbooks.Add(Type.Missing);
                Phase("after Workbooks.Add");
                sheets = workbook.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                ((Excel._Worksheet)sheet).Activate();
                SetValue(sheet, "A1", "original English source");
                SetValue(sheet, "B1", "second source cell");
                SetValue(sheet, "A2", "multicell selection neighbor");

                Phase("REAL Excel COM/menu/selection events; FAKE translation and controlled foreground signal. Native foreground acquisition is NOT asserted.");
                controller = new ExcelTranslator(excel, fakeTranslator, delegate { return simulatedExcelForeground; });
                menuTag = controller.MenuTag;
                string editMenuTag = controller.EditMenuTag;
                string nativeItemsBefore = NativeEditMenuItems(excel, editMenuTag);
                int nativeRequests = 0;
                controller.NativeTextTranslationRequested += delegate { nativeRequests++; };
                Phase("before ExcelTranslator.Start");
                controller.Start();
                Phase("after ExcelTranslator.Start");

                Check(NativeEditMenuItems(excel, editMenuTag) == nativeItemsBefore,
                    "adding translation preserves all original Formula Bar menu items and separators");
                Office.CommandBarControl editButton = FindButton(excel, editMenuTag, "Formula Bar");
                try
                {
                    Check(editButton != null, "native Formula Bar translation menu item is installed");
                    editButton.Execute();
                    WaitUntil(delegate { return nativeRequests == 1; }, 3000, "native text-menu callback");
                    Check(requests.Count == 0, "native edit-menu callback never falls back to whole-cell translation");
                }
                finally { Release(editButton); }

                Select(sheet, "A1");
                Office.CommandBarControl installed = FindButton(excel, menuTag);
                try { Check(installed != null, "real tagged cell context-menu button is installed"); }
                finally { Release(installed); }
                ExecuteButton(excel, controller, menuTag, requests, "first CommandBar Execute");
                try { WaitUntil(delegate { return requests.Count == 1 && controller.IsPopupVisible; }, 6000, "first translation popup"); }
                catch
                {
                    LogState("first popup timeout", excel, controller, requests);
                    throw;
                }
                Check(requests[0].Source == "original English source", "button Execute passes selected cell text to translator");
                Check(Convert.ToString(ReadValue(sheet, "A1")) == "original English source", "translation leaves original cell unchanged");
                Check(controller.LastSourceText == "original English source", "controller tracks the source displayed in the popup");
                CheckPopupAnchor(excel, controller, sheet, "A1", "100% zoom");

                Select(sheet, "B1");
                Check(!controller.IsPopupVisible, "selection-change event hides the active popup");

                delayNextRequest = true;
                ExecuteButton(excel, controller, menuTag, requests, "deferred CommandBar Execute");
                WaitUntil(delegate { return delayedRequest != null; }, 3000, "deferred translation request");
                Check(delayedRequest.Source == "second source cell", "deferred request captures the selected cell");
                Select(sheet, "A1");
                Check(!controller.IsPopupVisible, "selection change hides popup while a request is pending");
                delayedRequest.Completion.SetResult(new TranslationResult("late test translation", "test"));
                PumpFor(500);
                Check(!controller.IsPopupVisible, "late completion cannot reopen a popup for a stale selection");
                Check(Convert.ToString(ReadValue(sheet, "B1")) == "second source cell", "late translation leaves its source cell unchanged");

                delayNextRequest = true;
                delayedRequest = null;
                ExecuteButton(excel, controller, menuTag, requests, "focus-loss CommandBar Execute");
                WaitUntil(delegate { return delayedRequest != null && controller.IsPopupVisible; }, 3000, "focus-loss request");
                simulatedExcelForeground = false;
                PumpFor(400);
                Check(!controller.IsPopupVisible, "controlled foreground loss hides pending popup");
                delayedRequest.Completion.SetResult(new TranslationResult("late after focus loss", "test"));
                PumpFor(250);
                Check(!controller.IsPopupVisible, "late completion cannot reopen after controlled foreground loss");
                simulatedExcelForeground = true;

                mergedRange = sheet.get_Range("C1:D1", Type.Missing);
                mergedRange.Merge(Type.Missing);
                SetValue(sheet, "C1", "merged source text");
                Select(sheet, "C1");
                int beforeMerged = requests.Count;
                ExecuteButton(excel, controller, menuTag, requests, "merged-cell CommandBar Execute");
                WaitUntil(delegate { return requests.Count == beforeMerged + 1 && controller.IsPopupVisible; }, 6000, "merged-cell translation popup");
                Check(requests[beforeMerged].Source == "merged source text", "merged cell is accepted as one source cell");
                Check(Convert.ToString(ReadValue(sheet, "C1")) == "merged source text", "merged cell source remains unchanged");

                SetValue(sheet, "L15", "zoomed anchor source");
                Excel.Window activeWindow = excel.ActiveWindow;
                try
                {
                    activeWindow.Zoom = 150;
                    activeWindow.ScrollRow = 8;
                    activeWindow.ScrollColumn = 8;
                }
                finally { Release(activeWindow); }
                Select(sheet, "L15");
                int beforeZoomed = requests.Count;
                ExecuteButton(excel, controller, menuTag, requests, "zoomed-cell CommandBar Execute");
                WaitUntil(delegate { return requests.Count == beforeZoomed + 1 && controller.IsPopupVisible; }, 6000, "150% zoom scrolled-cell popup");
                Check(requests[beforeZoomed].Source == "zoomed anchor source", "scrolled cell text reaches translator");
                CheckPopupAnchor(excel, controller, sheet, "L15", "150% zoom with scrolling");

                Select(sheet, "A1:B1");
                Check(!controller.IsPopupVisible, "changing to a multi-cell selection hides the popup");
                int beforeRejected = requests.Count;
                ExecuteButton(excel, controller, menuTag, requests, "multi-cell CommandBar Execute");
                PumpFor(250);
                Check(requests.Count == beforeRejected, "multi-cell selection is rejected before translation");
                Check(Convert.ToString(ReadValue(sheet, "A1")) == "original English source" &&
                    Convert.ToString(ReadValue(sheet, "B1")) == "second source cell" &&
                    Convert.ToString(ReadValue(sheet, "C1")) == "merged source text" &&
                    Convert.ToString(ReadValue(sheet, "L15")) == "zoomed anchor source",
                    "all original cell contents remain unchanged");

                // Reproduce a COM callback arriving off the UI thread, followed by a
                // selection event before the UI queue is pumped. It must send no text.
                Select(sheet, "A1");
                int beforeQueuedRace = requests.Count;
                Exception queueError = null;
                var callbackThread = new Thread(delegate()
                {
                    try
                    {
                        var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                        typeof(ExcelTranslator).GetMethod("MenuButton_Click", flags).Invoke(controller, new object[] { null, false });
                        typeof(ExcelTranslator).GetMethod("Excel_SheetSelectionChange", flags).Invoke(controller, new object[] { null, null });
                    }
                    catch (Exception error) { queueError = error; }
                });
                callbackThread.IsBackground = true;
                callbackThread.Start();
                Check(callbackThread.Join(3000), "background Office callbacks enqueue without blocking the UI");
                if (queueError != null) throw queueError;
                PumpFor(150);
                Check(requests.Count == beforeQueuedRace, "selection event received before queued click executes prevents any translation request");

                int beforePartial = requests.Count;
                typeof(ExcelTranslator).GetMethod("TranslateSelectedText", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    .Invoke(controller, new object[] { "selected substring", new Point(400, 300) });
                WaitUntil(delegate { return requests.Count == beforePartial + 1 && controller.IsPopupVisible; }, 6000, "selected-text callback popup");
                Check(requests[beforePartial].Source == "selected substring", "selected-text callback never substitutes full cell value");
                Check((string)ReadValue(sheet, "A1") == "original English source", "selected-text translation leaves workbook unchanged");
                controller.ProcessSelectionChangeForTesting();
                Check(!controller.IsPopupVisible, "selected-text popup disappears on selection cancellation");
                delayNextRequest = true;
                delayedRequest = null;
                typeof(ExcelTranslator).GetMethod("TranslateSelectedText", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    .Invoke(controller, new object[] { "pending selected substring", new Point(400, 300) });
                WaitUntil(delegate { return delayedRequest != null; }, 3000, "delayed selected-text request");
                controller.ProcessSelectionChangeForTesting();
                delayedRequest.Completion.SetResult(new TranslationResult("late partial translation", "test"));
                PumpFor(200);
                Check(!controller.IsPopupVisible, "late selected-text result cannot reopen dismissed popup");

                controller.Dispose();
                controller = null;
                Office.CommandBarControl remaining = FindButton(excel, menuTag);
                try { Check(remaining == null, "controller cleanup removes its context-menu button"); }
                finally { Release(remaining); }
                Office.CommandBarControl remainingEdit = FindButton(excel, editMenuTag, "Formula Bar");
                try { Check(remainingEdit == null, "cleanup removes native edit translation item"); }
                finally { Release(remainingEdit); }
                Check(NativeEditMenuItems(excel, editMenuTag) == nativeItemsBefore,
                    "cleanup leaves original native edit menu intact");

                result = 0;
            }
            catch (Exception error)
            {
                Log.AppendLine(error.ToString());
                result = 1;
                WriteLog();
            }
            finally
            {
                Phase("cleanup controller");
                try
                {
                    if (controller != null) controller.Dispose();
                }
                catch (Exception error) { Log.AppendLine("Controller cleanup failed: " + error); result = 1; }
                Phase("cleanup temporary workbook");
                try
                {
                    if (workbook != null) workbook.Close(false, Type.Missing, Type.Missing);
                }
                catch (Exception error) { Log.AppendLine("Temporary workbook close failed: " + error); result = 1; }
                Phase("cleanup owned Excel instance");
                try
                {
                    if (excel != null) excel.Quit();
                }
                catch (Exception error) { Log.AppendLine("Private Excel instance quit failed: " + error); result = 1; }

                Release(mergedRange);
                Release(sheet);
                Release(sheets);
                Release(workbook);
                Release(workbooks);
                Release(excel);
                if (result == 0) Log.AppendLine("ALL PASSED: " + passed);
                else Log.AppendLine("FAILED after " + passed + " checks");
                WriteLog();
            }
            return result;
        }
    }
}
