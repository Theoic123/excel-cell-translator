using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Automation.Text;
using System.Windows.Forms;

namespace ExcelCellTranslator
{
    // A small, non-activating bridge for text selected inside Excel's edit controls.
    // Excel does not expose the active formula-bar / in-cell text range through its
    // object model, so this class uses UI Automation and an explicitly requested
    // Ctrl+C fallback. It never talks to Excel COM.
    internal sealed class EditSelectionBridge : IDisposable
    {
        private const int WhKeyboardLl = 13;
        private const int WhMouseLl = 14;
        private const int WmKeyDown = 0x0100;
        private const int WmKeyUp = 0x0101;
        private const int WmSysKeyDown = 0x0104;
        private const int WmSysKeyUp = 0x0105;
        private const int WmLButtonDown = 0x0201;
        private const int WmRButtonDown = 0x0204;
        private const int WmMButtonDown = 0x0207;
        private const int WmMouseWheel = 0x020A;
        private const int WmXButtonDown = 0x020B;
        private const int WmMouseHWheel = 0x020E;
        private const int VkControl = 0x11;
        private const int VkMenu = 0x12;
        private const int VkT = 0x54;
        private const int LlkhfInjected = 0x10;
        private const int KeyeventfKeyup = 0x0002;
        private const uint InputKeyboard = 1;
        private const uint GuiInMenuMode = 0x0004;
        private const uint GuiSystemMenuMode = 0x0008;
        private const uint GuiPopupMenuMode = 0x0010;
        private const int NativeMenuWaitAttempts = 30;
        private const int NativeMenuPollMilliseconds = 50;
        private const int MaximumSnapshotBytes = 16 * 1024 * 1024;
        private const long MaximumBitmapPixels = 16000000L;
        private static readonly UIntPtr CopyInputMarker = new UIntPtr(0x455843454C545241UL);

        private readonly Func<uint> _excelProcessId;
        private readonly Action<string, Point> _translate;
        private readonly Action _invalidate;
        private readonly Action<string, Point> _error;
        private readonly LowLevelKeyboardProc _keyboardProc;
        private readonly LowLevelMouseProc _mouseProc;
        private Control _dispatcher;
        private IntPtr _keyboardHook;
        private IntPtr _mouseHook;
        private int _uiThreadId;
        private bool _started;
        private bool _disposed;
        private bool _ownsLifecycle;
        private bool _candidateActive;
        private bool _hotkeyKeyupConsumed;
        private long _generation;
        private long _nativeMenuRequestId;
        private long _inputRevision;
        private long _uiaGeneration;
        private Task<string> _uiaTask;
        private System.Windows.Forms.Timer _focusTimer;
        private FocusSnapshot _focusSnapshot;
        private Point _anchor;
        private string _selectedText;

        internal EditSelectionBridge(Func<uint> excelProcessId, Action<string, Point> translate,
            Action invalidate, Action<string, Point> error)
        {
            if (excelProcessId == null) throw new ArgumentNullException("excelProcessId");
            if (translate == null) throw new ArgumentNullException("translate");
            if (invalidate == null) throw new ArgumentNullException("invalidate");
            if (error == null) throw new ArgumentNullException("error");
            _excelProcessId = excelProcessId;
            _translate = translate;
            _invalidate = invalidate;
            _error = error;
            _keyboardProc = KeyboardCallback;
            _mouseProc = MouseCallback;
        }

        // This property only reads native focus state. The caller can use it to
        // avoid Excel COM calls while Excel is editing text.
        internal bool IsEditing
        {
            get
            {
                FocusSnapshot snapshot;
                return TryCaptureFocus(out snapshot) && snapshot.RecognizedEdit;
            }
        }

        internal long InputRevision { get { return Interlocked.Read(ref _inputRevision); } }

        internal void Start()
        {
            if (_disposed) throw new ObjectDisposedException("EditSelectionBridge");
            if (_started) return;
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("EditSelectionBridge must be started on a Windows STA thread.");

            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            _dispatcher = new Control();
            IntPtr ignored = _dispatcher.Handle;
            _mouseHook = SetWindowsHookEx(WhMouseLl, _mouseProc, GetModuleHandle(null), 0);
            _keyboardHook = SetWindowsHookEx(WhKeyboardLl, _keyboardProc, GetModuleHandle(null), 0);
            _focusTimer = new System.Windows.Forms.Timer { Interval = 180 };
            _focusTimer.Tick += FocusTimer_Tick;
            _focusTimer.Start();
            _started = true;
            if (_mouseHook == IntPtr.Zero && _keyboardHook == IntPtr.Zero)
            {
                _ownsLifecycle = true;
                _error("无法监听 Excel 的文字编辑区，请重启工具后重试。", Cursor.Position);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Interlocked.Increment(ref _generation);
            if (_keyboardHook != IntPtr.Zero) { UnhookWindowsHookEx(_keyboardHook); _keyboardHook = IntPtr.Zero; }
            if (_mouseHook != IntPtr.Zero) { UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
            if (_focusTimer != null) { _focusTimer.Stop(); _focusTimer.Dispose(); _focusTimer = null; }
            if (_dispatcher != null) { _dispatcher.Dispose(); _dispatcher = null; }
            _uiaTask = null;
            _selectedText = null;
            // Shutdown intentionally does not call back into the parent controller.
        }

        // Public to this assembly for deterministic testing without reading the
        // user's clipboard or desktop focus. ValuePattern is deliberately absent:
        // it returns the whole control value rather than just the selected range.
        internal static string ReadAccessibleSelection(AutomationElement focusedElement, uint excelPid)
        {
            if (focusedElement == null || excelPid == 0) return null;
            try
            {
                AutomationElement element = focusedElement;
                for (int depth = 0; element != null && depth < 5; depth++)
                {
                    AutomationElement.AutomationElementInformation current = element.Current;
                    if (current.ProcessId != excelPid || current.IsPassword) return null;

                    object pattern;
                    if (element.TryGetCurrentPattern(TextPattern.Pattern, out pattern))
                    {
                        TextPattern textPattern = pattern as TextPattern;
                        if (textPattern == null) return null;
                        TextPatternRange[] ranges = textPattern.GetSelection();
                        if (ranges == null || ranges.Length != 1 || ranges[0] == null) return null;
                        string selected = ranges[0].GetText(TranslationService.MaximumCharacters + 1);
                        return string.IsNullOrEmpty(selected) ? null : selected;
                    }
                    element = TreeWalker.ControlViewWalker.GetParent(element);
                }
            }
            catch (ElementNotAvailableException) { }
            catch (InvalidOperationException) { }
            catch (COMException) { }
            catch (Exception) { /* Excel and third-party UIA providers can fail independently. */ }
            return null;
        }

        // Materialize the existing clipboard into known, self-contained data so
        // the explicit Ctrl+C fallback can restore it safely. Unknown OLE/rich
        // formats are refused before Ctrl+C is sent.
        internal static ClipboardSnapshot SnapshotClipboardData(IDataObject original)
        {
            var snapshot = new ClipboardSnapshot();
            if (original == null) return snapshot;

            string[] formats;
            try { formats = original.GetFormats(false); }
            catch (Exception ex) { snapshot.Dispose(); throw UnsafeClipboardError("无法读取", ex); }
            if (formats == null || formats.Length == 0) return snapshot;

            var data = new DataObject();
            try
            {
                for (int i = 0; i < formats.Length; i++)
                {
                    string format = formats[i];
                    object value = original.GetData(format, false);
                    if (value == null) throw UnsafeClipboardError("无法安全保存格式“" + format + "”", null);

                    string text = value as string;
                    if (text != null)
                    {
                        data.SetData(format, false, text);
                        continue;
                    }

                    string[] files = value as string[];
                    if (files != null)
                    {
                        data.SetData(format, false, (string[])files.Clone());
                        continue;
                    }

                    byte[] bytes = value as byte[];
                    if (bytes != null)
                    {
                        if (bytes.Length > MaximumSnapshotBytes)
                            throw UnsafeClipboardError("格式“" + format + "”过大", null);
                        data.SetData(format, false, (byte[])bytes.Clone());
                        continue;
                    }

                    MemoryStream memory = value as MemoryStream;
                    if (memory != null)
                    {
                        if (memory.Length > MaximumSnapshotBytes)
                            throw UnsafeClipboardError("格式“" + format + "”过大", null);
                        long oldPosition = memory.CanSeek ? memory.Position : 0;
                        byte[] content;
                        try
                        {
                            using (var copy = new MemoryStream())
                            {
                                if (memory.CanSeek) memory.Position = 0;
                                memory.CopyTo(copy);
                                content = copy.ToArray();
                            }
                        }
                        finally { if (memory.CanSeek) memory.Position = oldPosition; }
                        var clone = new MemoryStream(content, false);
                        snapshot._owned.Add(clone);
                        data.SetData(format, false, clone);
                        continue;
                    }

                    Bitmap bitmap = value as Bitmap;
                    if (bitmap != null)
                    {
                        if ((long)bitmap.Width * (long)bitmap.Height > MaximumBitmapPixels)
                            throw UnsafeClipboardError("剪贴板图片尺寸过大", null);
                        var clone = new Bitmap(bitmap);
                        snapshot._owned.Add(clone);
                        data.SetData(format, false, clone);
                        continue;
                    }

                    throw UnsafeClipboardError("无法安全保存格式“" + format + "”", null);
                }
            }
            catch
            {
                snapshot.Dispose();
                throw;
            }

            snapshot.DataObject = data;
            snapshot.IsEmpty = false;
            return snapshot;
        }

        private static InvalidOperationException UnsafeClipboardError(string reason, Exception inner)
        {
            string message = "剪贴板中有" + reason + "的内容，工具没有发送复制快捷键，也没有修改剪贴板。请先保存或移走该内容，再选中文字并重试。";
            return inner == null ? new InvalidOperationException(message) : new InvalidOperationException(message, inner);
        }

        private IntPtr KeyboardCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code < 0 || _disposed) return CallNextHookEx(_keyboardHook, code, wParam, lParam);
                int message = wParam.ToInt32();
                KeyboardHookData key = (KeyboardHookData)Marshal.PtrToStructure(lParam, typeof(KeyboardHookData));
                if ((key.Flags & LlkhfInjected) != 0 && key.ExtraInfo == CopyInputMarker)
                    return CallNextHookEx(_keyboardHook, code, wParam, lParam);

                if (_hotkeyKeyupConsumed && (message == WmKeyUp || message == WmSysKeyUp) && key.VirtualKey == VkT)
                {
                    _hotkeyKeyupConsumed = false;
                    return new IntPtr(1);
                }

                if ((message == WmKeyDown || message == WmSysKeyDown) && key.VirtualKey == VkT &&
                    (GetAsyncKeyState(VkControl) < 0) && (GetAsyncKeyState(VkMenu) < 0))
                {
                    FocusSnapshot snapshot;
                    if (TryCaptureFocus(out snapshot))
                    {
                        Interlocked.Increment(ref _nativeMenuRequestId);
                        Interlocked.Increment(ref _inputRevision);
                        _hotkeyKeyupConsumed = true;
                        DispatchOnUi(delegate { BeginHotkeyTranslation(snapshot, Cursor.Position); });
                        return new IntPtr(1);
                    }
                }

                if (message == WmKeyDown || message == WmSysKeyDown)
                {
                    Interlocked.Increment(ref _inputRevision);
                    DispatchOnUi(delegate { CancelForInput(); });
                }
                return CallNextHookEx(_keyboardHook, code, wParam, lParam);
            }
            catch (Exception) { return CallNextHookEx(_keyboardHook, code, wParam, lParam); }
        }

        private IntPtr MouseCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code >= 0 && !_disposed)
                {
                    int message = wParam.ToInt32();
                    if (message == WmLButtonDown || message == WmRButtonDown || message == WmMButtonDown ||
                        message == WmXButtonDown || message == WmMouseWheel || message == WmMouseHWheel)
                    {
                        Interlocked.Increment(ref _inputRevision);
                        DispatchOnUi(delegate { CancelForInput(); });
                    }
                }
            }
            catch (Exception) { }
            return CallNextHookEx(_mouseHook, code, wParam, lParam);
        }

        private void BeginHotkeyTranslation(FocusSnapshot snapshot, Point point)
        {
            if (_disposed || !SnapshotIsCurrent(snapshot)) return;
            EndPreviousLifecycle();
            long generation = Interlocked.Increment(ref _generation);
            _focusSnapshot = snapshot;
            _anchor = point;
            _selectedText = null;
            _candidateActive = true;
            _ownsLifecycle = true;
            Task<string> task = StartAccessibleRead(snapshot, generation);
            TranslateFromActionAsync(generation, task);
        }

        // Called by the native Formula Bar CommandBar Click handler. Excel may
        // still be closing its context menu and restoring focus to the editor,
        // so wait briefly for a recognized Excel edit control before reading.
        internal void TranslateSelectionFromNativeMenu(long expectedInputRevision)
        {
            long requestId = Interlocked.Increment(ref _nativeMenuRequestId);
            DispatchOnUi(delegate
            {
                BeginNativeMenuTranslation(expectedInputRevision, requestId, Interlocked.Read(ref _generation));
            });
        }

        private async void BeginNativeMenuTranslation(long expectedInputRevision, long requestId, long expectedGeneration)
        {
            // Let Excel unwind the native CommandBar callback and dismiss its popup
            // before probing the focus that belongs to the selected editor.
            await Task.Delay(NativeMenuPollMilliseconds);
            if (!NativeMenuWaitIsCurrent(expectedInputRevision, expectedGeneration, requestId)) return;

            FocusSnapshot snapshot = null;
            for (int i = 0; i < NativeMenuWaitAttempts; i++)
            {
                if (!NativeMenuWaitIsCurrent(expectedInputRevision, expectedGeneration, requestId)) return;
                if (TryCaptureFocus(out snapshot) && IsExcelMenuClosed(snapshot))
                {
                    if (!NativeMenuWaitIsCurrent(expectedInputRevision, expectedGeneration, requestId)) return;
                    if (!SnapshotIsCurrent(snapshot))
                    {
                        ShowNativeMenuTranslationError("无法确认 Excel 编辑区仍处于原来的位置，未读取或翻译内容。请重新选中文字后重试。");
                        return;
                    }
                    BeginHotkeyTranslation(snapshot, Cursor.Position);
                    return;
                }
                await Task.Delay(NativeMenuPollMilliseconds);
            }

            if (NativeMenuWaitIsCurrent(expectedInputRevision, expectedGeneration, requestId))
                ShowNativeMenuTranslationError("无法确认 Excel 公式栏或单元格文字编辑区已恢复焦点，未读取或翻译内容。请先关闭菜单并在编辑区选中文字后重试。");
        }

        private bool NativeMenuWaitIsCurrent(long expectedInputRevision, long expectedGeneration, long requestId)
        {
            return !_disposed && IsNativeMenuWaitCurrent(expectedInputRevision, InputRevision,
                expectedGeneration, Interlocked.Read(ref _generation), requestId,
                Interlocked.Read(ref _nativeMenuRequestId));
        }

        internal static bool IsNativeMenuWaitCurrent(long expectedInputRevision, long currentInputRevision,
            long expectedGeneration, long currentGeneration, long requestId, long currentRequestId)
        {
            return expectedInputRevision == currentInputRevision && expectedGeneration == currentGeneration &&
                requestId == currentRequestId;
        }

        private void ShowNativeMenuTranslationError(string message)
        {
            if (_disposed) return;
            _error(message, Cursor.Position);
        }

        private bool IsExcelMenuClosed(FocusSnapshot snapshot)
        {
            if (snapshot == null) return false;
            uint foregroundPid;
            uint thread = GetWindowThreadProcessId(snapshot.ForegroundWindow, out foregroundPid);
            if (thread == 0 || foregroundPid != snapshot.ProcessId) return false;
            var info = new GuiThreadInfo();
            info.Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo));
            if (!GetGUIThreadInfo(thread, ref info) || info.FocusWindow != snapshot.FocusWindow) return false;
            const uint menuFlags = GuiInMenuMode | GuiSystemMenuMode | GuiPopupMenuMode;
            return (info.Flags & menuFlags) == 0;
        }

        private async void TranslateFromActionAsync(long generation, Task<string> task)
        {
            try
            {
                // The hotkey or native menu command may still be unwinding.
                // Wait for modifier release so any fallback sends plain Ctrl+C.
                for (int i = 0; i < 24 && (GetAsyncKeyState(VkControl) < 0 || GetAsyncKeyState(VkMenu) < 0); i++)
                    await Task.Delay(25);

                if (!IsCurrentSession(generation)) return;
                string text = _selectedText;
                if (string.IsNullOrEmpty(text) && task != null)
                {
                    Task completed = await Task.WhenAny(task, Task.Delay(800));
                    if (object.ReferenceEquals(completed, task))
                    {
                        try { text = await task; } catch { text = null; }
                    }
                }

                if (!IsCurrentSession(generation)) return;
                if (string.IsNullOrEmpty(text))
                {
                    if (_focusSnapshot == null || !_focusSnapshot.RecognizedEdit)
                    {
                        _error("Excel 没有提供所选文字。请在单元格文字编辑区或公式栏中选中文字后重试。", _anchor);
                        return;
                    }
                    text = await CopySelectedTextAsync(generation, _focusSnapshot);
                }

                if (!IsCurrentSession(generation)) return;
                if (string.IsNullOrWhiteSpace(text))
                {
                    _error("没有读到所选文字。请先在单元格文字编辑区或公式栏中选中文字，再选择“翻译所选文字”。", _anchor);
                    return;
                }

                _candidateActive = false;
                _ownsLifecycle = true;
                _translate(text, _anchor);
            }
            catch (Exception ex)
            {
                if (IsCurrentSession(generation))
                    _error(string.IsNullOrWhiteSpace(ex.Message) ? "无法读取所选文字，请重新选中后重试。" : ex.Message, _anchor);
            }
        }

        private async Task<string> CopySelectedTextAsync(long generation, FocusSnapshot snapshot)
        {
            if (!IsCurrentSession(generation) || !snapshot.RecognizedEdit)
                throw new InvalidOperationException("Excel 编辑区已改变，请重新选中文字后重试。");
            if (!CopyModifiersReleased())
                throw new InvalidOperationException("请先松开 Ctrl、Alt 和 Shift，再重新选择翻译所选文字。");

            uint before = GetClipboardSequenceNumber();
            ClipboardSnapshot previous;
            try { previous = SnapshotClipboardData(Clipboard.GetDataObject()); }
            catch (InvalidOperationException) { throw; }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法安全读取当前剪贴板，因此没有发送复制快捷键。请关闭正在占用剪贴板的程序后重试。", ex);
            }

            using (previous)
            {
                if (GetClipboardSequenceNumber() != before)
                    throw new InvalidOperationException("剪贴板在读取时发生变化，工具没有覆盖新内容。请确认后重试。");
                if (!IsCurrentSession(generation) || !SnapshotIsCurrent(snapshot))
                    throw new InvalidOperationException("Excel 编辑区已改变，未复制文字。请重新选中文字后重试。");
                if (!CopyModifiersReleased())
                    throw new InvalidOperationException("请先松开 Ctrl、Alt 和 Shift，再重新选择翻译所选文字。");
                long inputRevisionAtCopy = Interlocked.Read(ref _inputRevision);
                if (GetClipboardSequenceNumber() != before)
                    throw new InvalidOperationException("剪贴板在发送复制快捷键前发生变化，工具没有覆盖新内容。请确认后重试。");
                if (!SendCopyInput()) throw new InvalidOperationException("无法发送复制快捷键。请重新选中文字后重试。");

                uint copiedSequence = 0;
                string copiedText = null;
                bool ownsClipboardUpdate = false;
                bool stableCopyRead = false;
                try
                {
                    for (int i = 0; i < 20; i++)
                    {
                        await Task.Delay(60);
                        if (!IsCurrentSession(generation) || !SnapshotIsCurrent(snapshot) ||
                            Interlocked.Read(ref _inputRevision) != inputRevisionAtCopy)
                        {
                            CaptureOwnedClipboardUpdate(before, inputRevisionAtCopy, snapshot.ProcessId,
                                ref copiedSequence, ref ownsClipboardUpdate);
                            throw new InvalidOperationException("读取期间 Excel 编辑区已改变，未使用复制结果。请重新选中文字后重试。");
                        }
                        uint after = GetClipboardSequenceNumber();
                        if (after == before) continue;
                        uint ownerPid;
                        if (!ClipboardOwnerBelongsTo(snapshot.ProcessId, out ownerPid))
                            throw new InvalidOperationException("剪贴板在读取时被其他程序更新，工具没有读取或恢复该内容。请重试。");
                        copiedSequence = after;
                        ownsClipboardUpdate = true;
                        string candidate = Clipboard.ContainsText() ? Clipboard.GetText(TextDataFormat.UnicodeText) : "";
                        uint stable = GetClipboardSequenceNumber();
                        uint stableOwnerPid;
                        if (stable != after || !ClipboardOwnerBelongsTo(snapshot.ProcessId, out stableOwnerPid) || stableOwnerPid != snapshot.ProcessId)
                            throw new InvalidOperationException("剪贴板在读取时再次变化，工具没有使用该内容。请重试。");
                        copiedSequence = stable;
                        copiedText = candidate;
                        stableCopyRead = true;
                        break;
                    }

                    if (!stableCopyRead)
                        throw new InvalidOperationException("Excel 没有复制所选文字。请先确认文字已选中，再重试。");
                    return copiedText ?? "";
                }
                catch (ExternalException ex)
                {
                    throw new InvalidOperationException("Excel 复制了文字，但无法读取剪贴板。请重试。", ex);
                }
                finally
                {
                    // Restoration is tied to the exact, stable Excel clipboard
                    // update. It does not depend on focus still being in Excel.
                    if (ownsClipboardUpdate)
                        TryRestoreClipboard(previous, copiedSequence, snapshot.ProcessId);
                }
            }
        }

        private static bool ClipboardOwnerBelongsTo(uint expectedPid, out uint ownerPid)
        {
            ownerPid = 0;
            IntPtr owner = GetClipboardOwner();
            return owner != IntPtr.Zero && GetWindowThreadProcessId(owner, out ownerPid) != 0 && ownerPid == expectedPid;
        }

        private void CaptureOwnedClipboardUpdate(uint before, long inputRevisionAtCopy, uint expectedPid,
            ref uint copiedSequence, ref bool ownsUpdate)
        {
            uint current = GetClipboardSequenceNumber();
            uint ownerPid;
            if (current != before && ClipboardOwnerBelongsTo(expectedPid, out ownerPid) &&
                CanClaimClipboardUpdate(before, current, inputRevisionAtCopy,
                    Interlocked.Read(ref _inputRevision), ownerPid, expectedPid))
            {
                uint stable = GetClipboardSequenceNumber();
                uint stableOwner;
                long stableInputRevision = Interlocked.Read(ref _inputRevision);
                if (stable == current && ClipboardOwnerBelongsTo(expectedPid, out stableOwner) &&
                    CanClaimClipboardUpdate(before, stable, inputRevisionAtCopy, stableInputRevision, stableOwner, expectedPid))
                {
                    copiedSequence = stable;
                    ownsUpdate = true;
                }
            }
        }

        internal static bool CanClaimClipboardUpdate(uint beforeSequence, uint currentSequence,
            long beforeInputRevision, long currentInputRevision, uint clipboardOwnerProcessId, uint expectedProcessId)
        {
            return expectedProcessId != 0 && currentSequence != beforeSequence &&
                beforeInputRevision == currentInputRevision && clipboardOwnerProcessId == expectedProcessId;
        }

        private static void TryRestoreClipboard(ClipboardSnapshot previous, uint expectedSequence, uint expectedPid)
        {
            uint ownerPid;
            if (expectedSequence == 0 || GetClipboardSequenceNumber() != expectedSequence ||
                !ClipboardOwnerBelongsTo(expectedPid, out ownerPid) || ownerPid != expectedPid) return;
            try
            {
                if (previous.IsEmpty) Clipboard.Clear();
                else Clipboard.SetDataObject(previous.DataObject, true);
            }
            catch (Exception) { /* Do not overwrite later user clipboard changes. */ }
        }

        private Task<string> StartAccessibleRead(FocusSnapshot snapshot, long generation)
        {
            if (_uiaTask != null && !_uiaTask.IsCompleted) return null;
            _uiaGeneration = generation;
            Task<string> task = Task.Run(delegate
            {
                if (!SnapshotIsCurrent(snapshot)) return null;
                AutomationElement nativeFocused = null;
                try { nativeFocused = AutomationElement.FromHandle(snapshot.FocusWindow); }
                catch { }
                if (!SnapshotIsCurrent(snapshot)) return null;
                bool nativeIsEdit = nativeFocused != null &&
                    (snapshot.RecognizedEdit || HasEditControlType(nativeFocused, snapshot.ProcessId));
                if (nativeIsEdit)
                {
                    string nativeSelection = ReadAccessibleSelection(nativeFocused, snapshot.ProcessId);
                    if (!SnapshotIsCurrent(snapshot)) return null;
                    if (!string.IsNullOrEmpty(nativeSelection)) return nativeSelection;
                }

                AutomationElement focused = null;
                try { focused = AutomationElement.FocusedElement; }
                catch { }
                if (!SnapshotIsCurrent(snapshot)) return null;
                bool globalIsEdit = focused != null &&
                    (snapshot.RecognizedEdit || HasEditControlType(focused, snapshot.ProcessId));
                if (!globalIsEdit || object.ReferenceEquals(focused, nativeFocused)) return null;
                string selected = ReadAccessibleSelection(focused, snapshot.ProcessId);
                if (!SnapshotIsCurrent(snapshot)) return null;
                return selected;
            });
            _uiaTask = task;
            task.ContinueWith(delegate(Task<string> completed)
            {
                string text = null;
                if (completed.Status == TaskStatus.RanToCompletion) text = completed.Result;
                DispatchOnUi(delegate { AccessibleReadCompleted(generation, snapshot, task, text); });
            }, TaskScheduler.Default);
            return task;
        }

        private void AccessibleReadCompleted(long generation, FocusSnapshot snapshot, Task<string> task, string text)
        {
            if (_disposed) return;
            if (_uiaGeneration == generation && object.ReferenceEquals(_uiaTask, task)) _uiaTask = null;
            if (Interlocked.Read(ref _generation) != generation || !_candidateActive || !SnapshotIsCurrent(snapshot)) return;
            _selectedText = text;
        }

        private void CancelForInput()
        {
            if (_disposed) return;
            if (_ownsLifecycle) EndPreviousLifecycle();
            else if (_candidateActive)
            {
                _candidateActive = false;
                _selectedText = null;
                Interlocked.Increment(ref _generation);
            }
        }

        private void EndPreviousLifecycle()
        {
            bool notify = _ownsLifecycle;
            _ownsLifecycle = false;
            _candidateActive = false;
            _selectedText = null;
            _focusSnapshot = null;
            if (notify && !_disposed) _invalidate();
        }

        private bool IsCurrentSession(long generation)
        {
            return !_disposed && Interlocked.Read(ref _generation) == generation && _focusSnapshot != null && SnapshotIsCurrent(_focusSnapshot);
        }

        private bool TryCaptureFocus(out FocusSnapshot snapshot)
        {
            snapshot = null;
            uint expected = _excelProcessId();
            if (expected == 0) return false;
            IntPtr foreground = GetForegroundWindow();
            uint foregroundPid;
            uint thread = GetWindowThreadProcessId(foreground, out foregroundPid);
            if (foreground == IntPtr.Zero || foregroundPid != expected || thread == 0) return false;

            var info = new GuiThreadInfo();
            info.Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo));
            if (!GetGUIThreadInfo(thread, ref info)) return false;
            IntPtr focus = info.FocusWindow;
            uint focusPid;
            if (focus == IntPtr.Zero || GetWindowThreadProcessId(focus, out focusPid) == 0 || focusPid != expected) return false;

            bool recognizedFocus = IsRecognizedEditWindow(focus, expected);
            snapshot = new FocusSnapshot(foreground, focus, expected, recognizedFocus, Cursor.Position);
            return true;
        }

        private void FocusTimer_Tick(object sender, EventArgs e)
        {
            if ((_candidateActive || _ownsLifecycle) && _focusSnapshot != null && !SnapshotIsCurrent(_focusSnapshot))
                CancelForInput();
        }

        private static bool HasEditControlType(AutomationElement element, uint processId)
        {
            try
            {
                AutomationElement current = element;
                for (int depth = 0; current != null && depth < 4; depth++)
                {
                    AutomationElement.AutomationElementInformation info = current.Current;
                    if (info.ProcessId != processId) return false;
                    if (info.ControlType == ControlType.Edit) return true;
                    current = TreeWalker.ControlViewWalker.GetParent(current);
                }
            }
            catch { }
            return false;
        }

        private static bool IsRecognizedEditWindow(IntPtr window, uint expectedPid)
        {
            IntPtr current = window;
            for (int depth = 0; current != IntPtr.Zero && depth < 6; depth++)
            {
                uint pid;
                if (GetWindowThreadProcessId(current, out pid) == 0 || pid != expectedPid) break;
                string className = GetClassNameOf(current);
                if (string.Equals(className, "EXCEL6", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(className, "EXCEL<EDIT>", StringComparison.OrdinalIgnoreCase)) return true;
                current = GetParent(current);
            }
            return false;
        }

        private bool SnapshotIsCurrent(FocusSnapshot snapshot)
        {
            if (snapshot == null || _disposed || _excelProcessId() != snapshot.ProcessId) return false;
            if (GetForegroundWindow() != snapshot.ForegroundWindow) return false;
            uint foregroundPid;
            uint thread = GetWindowThreadProcessId(snapshot.ForegroundWindow, out foregroundPid);
            if (foregroundPid != snapshot.ProcessId || thread == 0) return false;
            var info = new GuiThreadInfo();
            info.Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo));
            if (!GetGUIThreadInfo(thread, ref info) || info.FocusWindow != snapshot.FocusWindow) return false;
            uint focusPid;
            return GetWindowThreadProcessId(info.FocusWindow, out focusPid) != 0 && focusPid == snapshot.ProcessId;
        }

        private void DispatchOnUi(Action action)
        {
            if (_disposed || action == null) return;
            if (Thread.CurrentThread.ManagedThreadId == _uiThreadId)
            {
                action();
                return;
            }
            Control dispatcher = _dispatcher;
            if (dispatcher == null || dispatcher.IsDisposed || !dispatcher.IsHandleCreated) return;
            try { dispatcher.BeginInvoke(action); } catch (InvalidOperationException) { }
        }

        private static bool SendCopyInput()
        {
            var input = new Input[4];
            ushort[] keys = { 0x11, 0x43, 0x43, 0x11 };
            for (int i = 0; i < input.Length; i++)
            {
                input[i].Type = InputKeyboard;
                input[i].Data.Keyboard.VirtualKey = keys[i];
                input[i].Data.Keyboard.Flags = i >= 2 ? (uint)KeyeventfKeyup : 0U;
                input[i].Data.Keyboard.ExtraInfo = CopyInputMarker;
            }
            return SendInput((uint)input.Length, input, Marshal.SizeOf(typeof(Input))) == input.Length;
        }

        private static bool CopyModifiersReleased()
        {
            return GetAsyncKeyState(0x10) >= 0 && GetAsyncKeyState(VkControl) >= 0 && GetAsyncKeyState(VkMenu) >= 0 &&
                GetAsyncKeyState(0x5B) >= 0 && GetAsyncKeyState(0x5C) >= 0;
        }

        private sealed class FocusSnapshot
        {
            internal readonly IntPtr ForegroundWindow;
            internal readonly IntPtr FocusWindow;
            internal readonly uint ProcessId;
            internal readonly bool RecognizedEdit;
            internal readonly Point Point;

            internal FocusSnapshot(IntPtr foreground, IntPtr focus, uint pid, bool recognizedEdit, Point point)
            {
                ForegroundWindow = foreground;
                FocusWindow = focus;
                ProcessId = pid;
                RecognizedEdit = recognizedEdit;
                Point = point;
            }
        }

        internal sealed class ClipboardSnapshot : IDisposable
        {
            private bool _disposed;
            internal IDataObject DataObject { get; set; }
            internal bool IsEmpty { get; set; }
            internal readonly List<IDisposable> _owned = new List<IDisposable>();

            internal ClipboardSnapshot() { IsEmpty = true; }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                for (int i = _owned.Count - 1; i >= 0; i--)
                {
                    try { _owned[i].Dispose(); } catch { }
                }
                _owned.Clear();
                DataObject = null;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardHookData
        {
            internal uint VirtualKey;
            internal uint ScanCode;
            internal uint Flags;
            internal uint Time;
            internal UIntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            internal uint Size;
            internal uint Flags;
            internal IntPtr ActiveWindow;
            internal IntPtr FocusWindow;
            internal IntPtr CaptureWindow;
            internal IntPtr MenuOwnerWindow;
            internal IntPtr MoveSizeWindow;
            internal IntPtr CaretWindow;
            internal NativeRect CaretRect;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            internal int Left;
            internal int Top;
            internal int Right;
            internal int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Input
        {
            internal uint Type;
            internal InputUnion Data;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] internal KeyboardInput Keyboard;
            // INPUT contains the full Win32 union even when we only send keys.
            // MOUSEINPUT sets its size/alignment (40 bytes for INPUT on x64).
            [FieldOffset(0)] internal MouseInput Mouse;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MouseInput
        {
            internal int X;
            internal int Y;
            internal uint MouseData;
            internal uint Flags;
            internal uint Time;
            internal UIntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardInput
        {
            internal ushort VirtualKey;
            internal ushort ScanCode;
            internal uint Flags;
            internal uint Time;
            internal UIntPtr ExtraInfo;
        }

        private delegate IntPtr LowLevelKeyboardProc(int code, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelMouseProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc callback, IntPtr module, uint threadId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc callback, IntPtr module, uint threadId);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo info);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetClassName(IntPtr window, System.Text.StringBuilder className, int maxCount);
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint count, Input[] inputs, int size);
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int virtualKey);
        [DllImport("user32.dll")]
        private static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll")]
        private static extern IntPtr GetClipboardOwner();

        private static string GetClassNameOf(IntPtr window)
        {
            if (window == IntPtr.Zero) return "";
            var name = new System.Text.StringBuilder(128);
            return GetClassName(window, name, name.Capacity) > 0 ? name.ToString() : "";
        }
    }
}
