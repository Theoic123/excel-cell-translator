using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelCellTranslator
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (HasArgument(args, "--self-test")) return SelfTests.Run();
            if (HasArgument(args, "--excel-test")) return SelfTests.RunExcel();

            bool createdNew;
            using (var mutex = new Mutex(true, @"Local\ExcelCellTranslator.Singleton", out createdNew))
            {
                if (!createdNew)
                {
                    MessageBox.Show("Excel 翻译工具已经在运行中。请从系统托盘打开它。", "Excel 翻译工具",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return 0;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using (var context = new TranslatorApplicationContext()) Application.Run(context);
            }
            return 0;
        }

        private static bool HasArgument(string[] args, string expected)
        {
            if (args == null) return false;
            foreach (string arg in args)
                if (string.Equals(arg, expected, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
    }

    internal sealed class TranslatorApplicationContext : ApplicationContext
    {
        private readonly MainForm _mainForm;
        private readonly ExcelTranslator _translator;
        private readonly NotifyIcon _trayIcon;
        private bool _allowExit;
        private bool _hideToTray = true;

        internal TranslatorApplicationContext()
        {
            _mainForm = new MainForm();
            _translator = new ExcelTranslator();
            _mainForm.SettingsRequested += MainForm_SettingsRequested;
            _mainForm.InfoRequested += MainForm_InfoRequested;
            _mainForm.ExitRequested += MainForm_ExitRequested;
            _mainForm.CloseRequested += MainForm_CloseRequested;
            _mainForm.HideToTray = _hideToTray;
            _translator.StatusChanged += Translator_StatusChanged;
            _mainForm.SetStatus(_translator.StatusText);

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("打开主窗口", null, delegate { ShowMainForm(); });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("退出", null, delegate { ExitApplication(); });
            _trayIcon = new NotifyIcon
            {
                Text = "Excel 翻译工具",
                Icon = AppIcon.Create(),
                ContextMenuStrip = trayMenu,
                Visible = true
            };
            _trayIcon.DoubleClick += delegate { ShowMainForm(); };

            _mainForm.Show();
            _translator.Start();
        }

        private void Translator_StatusChanged(object sender, EventArgs e)
        {
            _mainForm.SetStatus(_translator.StatusText);
        }

        private void MainForm_SettingsRequested(object sender, EventArgs e)
        {
            using (var settings = new SettingsForm(_hideToTray, _translator.ShowProgressPopup))
            {
                if (settings.ShowDialog(_mainForm) == DialogResult.OK)
                {
                    _hideToTray = settings.HideToTray;
                    _mainForm.HideToTray = _hideToTray;
                    _translator.ShowProgressPopup = settings.ShowProgressPopup;
                }
            }
        }

        private void MainForm_InfoRequested(object sender, EventArgs e)
        {
            using (var info = new InfoForm()) info.ShowDialog(_mainForm);
        }

        private void MainForm_ExitRequested(object sender, EventArgs e)
        {
            ExitApplication();
        }

        private void MainForm_CloseRequested(object sender, FormClosingEventArgs e)
        {
            if (_allowExit) return;
            if (_hideToTray)
            {
                e.Cancel = true;
                _mainForm.Hide();
            }
            else ExitApplication();
        }

        private void ShowMainForm()
        {
            if (_mainForm.IsDisposed) return;
            if (!_mainForm.Visible) _mainForm.Show();
            if (_mainForm.WindowState == FormWindowState.Minimized) _mainForm.WindowState = FormWindowState.Normal;
            _mainForm.Activate();
        }

        private void ExitApplication()
        {
            if (_allowExit) return;
            _allowExit = true;
            _trayIcon.Visible = false;
            _translator.Dispose();
            _mainForm.Close();
            ExitThread();
        }

        protected override void ExitThreadCore()
        {
            _allowExit = true;
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _translator.Dispose();
            if (!_mainForm.IsDisposed) _mainForm.Dispose();
            base.ExitThreadCore();
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly Label _statusLabel;
        private readonly Label _statusDot;
        private readonly Panel _statusCard;

        internal event EventHandler SettingsRequested;
        internal event EventHandler InfoRequested;
        internal event EventHandler ExitRequested;
        internal event FormClosingEventHandler CloseRequested;
        internal bool HideToTray { get; set; }

        internal MainForm()
        {
            Text = "Excel 翻译工具";
            MinimumSize = new Size(440, 390);
            ClientSize = new Size(490, 450);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(246, 248, 252);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var title = new Label
            {
                Text = "Excel 翻译工具",
                Font = new Font("Microsoft YaHei UI", 18F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 43, 66),
                Location = new Point(26, 22),
                AutoSize = true
            };
            var subtitle = new Label
            {
                Text = "选中单元格后右键即可翻译",
                ForeColor = Color.FromArgb(107, 119, 139),
                Location = new Point(28, 70),
                AutoSize = true
            };

            _statusCard = new Panel
            {
                Location = new Point(24, 108),
                Size = new Size(436, 70),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            _statusDot = new Label
            {
                Text = "●",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(224, 155, 57),
                Location = new Point(18, 20)
            };
            _statusLabel = new Label
            {
                Text = "正在启动…",
                AutoSize = false,
                Location = new Point(48, 14),
                Size = new Size(365, 42),
                ForeColor = Color.FromArgb(43, 55, 77),
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _statusCard.Controls.Add(_statusDot);
            _statusCard.Controls.Add(_statusLabel);

            var howTitle = SectionTitle("使用方法", 190);
            var howText = BodyText("选中单元格或编辑区文字，右键选择翻译。\r\n局部选文也可按 Ctrl+Alt+T。\r\n中英自动切换；切换选区后浮窗消失。", 218, 74);
            var privacyTitle = SectionTitle("联网说明", 302);
            var privacyText = BodyText("触发翻译后，文字会发送到所选服务。\r\n局部选文兼容读取可能临时使用剪贴板。", 330, 56);

            var settingsButton = MakeButton("设置", 24, 398, 100);
            settingsButton.Click += delegate { if (SettingsRequested != null) SettingsRequested(this, EventArgs.Empty); };
            var infoButton = MakeButton("关于与说明", 134, 398, 120);
            infoButton.Click += delegate { if (InfoRequested != null) InfoRequested(this, EventArgs.Empty); };
            var exitButton = MakeButton("退出", 360, 398, 100);
            exitButton.Click += delegate { if (ExitRequested != null) ExitRequested(this, EventArgs.Empty); };

            Controls.Add(title);
            Controls.Add(subtitle);
            Controls.Add(_statusCard);
            Controls.Add(howTitle);
            Controls.Add(howText);
            Controls.Add(privacyTitle);
            Controls.Add(privacyText);
            Controls.Add(settingsButton);
            Controls.Add(infoButton);
            Controls.Add(exitButton);
        }

        private static Label SectionTitle(string text, int top)
        {
            return new Label
            {
                Text = text,
                Location = new Point(26, top),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(43, 55, 77)
            };
        }

        private static Label BodyText(string text, int top, int height)
        {
            return new Label
            {
                Text = text,
                Location = new Point(28, top),
                Size = new Size(430, height),
                ForeColor = Color.FromArgb(98, 109, 128),
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular),
                AutoEllipsis = true
            };
        }

        private static Button MakeButton(string text, int left, int top, int width)
        {
            return new Button
            {
                Text = text,
                Location = new Point(left, top),
                Size = new Size(width, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(48, 83, 139),
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
        }

        internal void SetStatus(string text)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action<string>(SetStatus), text); } catch (InvalidOperationException) { }
                return;
            }
            _statusLabel.Text = text ?? string.Empty;
            bool connected = text != null && text.IndexOf("已连接", StringComparison.Ordinal) >= 0;
            _statusDot.ForeColor = connected ? Color.FromArgb(46, 164, 111) : Color.FromArgb(224, 155, 57);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (CloseRequested != null) CloseRequested(this, e);
            base.OnFormClosing(e);
        }
    }

    internal sealed class SettingsForm : Form
    {
        private readonly CheckBox _hideToTray;
        private readonly CheckBox _showProgress;
        internal bool HideToTray { get { return _hideToTray.Checked; } }
        internal bool ShowProgressPopup { get { return _showProgress.Checked; } }

        internal SettingsForm(bool hideToTray, bool showProgress)
        {
            Text = "设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(390, 265);
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(248, 250, 253);
            var intro = new Label { Text = "窗口偏好本次生效；翻译服务设置单独保存。", Location = new Point(22, 18), AutoSize = true, ForeColor = Color.FromArgb(99, 111, 130) };
            _hideToTray = new CheckBox { Text = "关闭主窗口时收起到系统托盘", Location = new Point(24, 58), AutoSize = true, Checked = hideToTray };
            _showProgress = new CheckBox { Text = "翻译时显示“正在翻译”提示", Location = new Point(24, 91), AutoSize = true, Checked = showProgress };
            var note = new Label
            {
                Text = "Excel 连接会自动重试；本工具只接入一个 Excel 实例。",
                Location = new Point(24, 128), Size = new Size(340, 32),
                ForeColor = Color.FromArgb(99, 111, 130)
            };
            var save = new Button { Text = "保存", DialogResult = DialogResult.OK, Location = new Point(274, 220), Size = new Size(90, 29) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(174, 220), Size = new Size(90, 29) };
            var services = new Button { Text = "翻译服务 / API 密钥", Location = new Point(24, 168), Size = new Size(220, 32) };
            services.Click += delegate { using (var form = new ProviderSettingsForm()) form.ShowDialog(this); };
            Controls.Add(services); Controls.Add(intro); Controls.Add(_hideToTray); Controls.Add(_showProgress); Controls.Add(note); Controls.Add(save); Controls.Add(cancel);
            AcceptButton = save; CancelButton = cancel;
        }
    }

    internal sealed class InfoForm : Form
    {
        internal InfoForm()
        {
            Text = "关于与联网说明";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(440, 255);
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(248, 250, 253);
            var title = new Label { Text = "Excel 翻译工具", Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold), Location = new Point(22, 18), AutoSize = true, ForeColor = Color.FromArgb(31, 43, 66) };
            var text = new Label
            {
                Text = "可在设置中选择 MyMemory、Google 或 DeepL。只有触发翻译后，单元格内容或选中的文字才会发送给所选服务。\r\n\r\n" +
                       "编辑区优先通过辅助功能读取选文；兼容方式会复制选文并恢复可安全保留的剪贴板。不修改工作簿内容。\r\n\r\n" +
                       "如果同时打开多个独立 Excel 实例，本工具只自动连接到其中一个。",
                Location = new Point(24, 62), Size = new Size(390, 142),
                ForeColor = Color.FromArgb(76, 89, 110)
            };
            var close = new Button { Text = "关闭", DialogResult = DialogResult.OK, Location = new Point(324, 206), Size = new Size(90, 30) };
            Controls.Add(title); Controls.Add(text); Controls.Add(close); AcceptButton = close;
        }
    }

    internal sealed class ExcelTranslator : IDisposable
    {
        private const string MenuPrefix = "ExcelCellTranslator.";
        private const string MenuCaption = "翻译（中英自动切换）";
        private const int ExcelRetryMilliseconds = 1500;
        private const int ForegroundPollMilliseconds = 180;
        private readonly Excel.Application _injectedExcel;
        private readonly Func<string, CancellationToken, Task<TranslationResult>> _translate;
        private readonly Func<bool> _foregroundCheck;
        private readonly string _menuTag = MenuPrefix + Guid.NewGuid().ToString("N");
        private string _tableMenuTag;
        private Excel.Application _excel;
        private bool _ownsExcelReference;
        private bool _started;
        private bool _disposed;
        private bool _selectionEventAttached;
        private bool _rightClickEventAttached;
        private bool _sheetActivateEventAttached;
        private bool _windowDeactivateEventAttached;
        private bool _staleTagsCleaned;
        private int _excelHwnd;
        private uint _excelProcessId;
        private Office.CommandBarButton _cellButton;
        private Office.CommandBarButton _listButton;
        private Office.CommandBarButton _editButton;
        private System.Windows.Forms.Timer _attachTimer;
        private System.Windows.Forms.Timer _foregroundTimer;
        private TranslationPopupForm _popup;
        private Point _popupAnchor = Point.Empty;
        private string _statusText = "正在查找已打开的 Excel…";
        private CancellationTokenSource _activeCancellation;
        private long _generation;
        private long _eventRevision;
        private string _lastSourceText;
        private bool _showProgressPopup = true;
        private GlobalDismissHooks _dismissHooks;
        private EditSelectionBridge _editSelection;
        private Control _dispatcher;
        private int _uiThreadId;

        internal event EventHandler StatusChanged;
        internal string StatusText { get { return _statusText; } }
        internal bool ShowProgressPopup { get { return _showProgressPopup; } set { _showProgressPopup = value; } }
        internal string MenuTag { get { return _menuTag; } }
        internal string EditMenuTag { get { return _menuTag + ".EditText"; } }
        internal event Action NativeTextTranslationRequested;
        internal bool IsPopupVisible { get { return _popup != null && !_popup.IsDisposed && _popup.Visible; } }
        internal Rectangle PopupBounds { get { return IsPopupVisible ? _popup.Bounds : Rectangle.Empty; } }
        internal Point PopupAnchor { get { return _popupAnchor; } }
        internal long CurrentGeneration { get { return _generation; } }
        internal string LastSourceText { get { return _lastSourceText; } }

        internal ExcelTranslator()
            : this(null, null)
        {
        }

        internal ExcelTranslator(Excel.Application excel, Func<string, CancellationToken, Task<TranslationResult>> translator)
            : this(excel, translator, null)
        {
        }

        // Tests may control this external signal without stealing the user's desktop focus.
        // Normal startup always uses the native foreground-window/PID check below.
        internal ExcelTranslator(Excel.Application excel, Func<string, CancellationToken, Task<TranslationResult>> translator, Func<bool> foregroundCheck)
        {
            _injectedExcel = excel;
            _translate = translator ?? TranslationService.TranslateAsync;
            _foregroundCheck = foregroundCheck;
            _tableMenuTag = _menuTag + ".Table";
        }

        internal void Start()
        {
            if (_disposed) throw new ObjectDisposedException("ExcelTranslator");
            if (_started) return;
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("ExcelTranslator must be started on a Windows STA thread.");
            _started = true;
            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            _dispatcher = new Control();
            IntPtr dispatcherHandle = _dispatcher.Handle;
            _foregroundTimer = new System.Windows.Forms.Timer { Interval = ForegroundPollMilliseconds };
            _foregroundTimer.Tick += ForegroundTimer_Tick;
            _foregroundTimer.Start();
            if (_foregroundCheck == null)
            {
                _editSelection = new EditSelectionBridge(() => _excelProcessId, TranslateSelectedText, InvalidateCurrentPopup,
                    (message, point) => ShowPopup("无法读取所选文字", message, false, point));
                _editSelection.Start();
            }

            if (_injectedExcel != null)
            {
                AttachToExcel(_injectedExcel, false);
                return;
            }

            _attachTimer = new System.Windows.Forms.Timer { Interval = ExcelRetryMilliseconds };
            _attachTimer.Tick += AttachTimer_Tick;
            _attachTimer.Start();
            TryAttachToRunningExcel();
        }

        private void AttachTimer_Tick(object sender, EventArgs e)
        {
            if (_editSelection != null && _editSelection.IsEditing) return;
            if (_excel == null)
            {
                if (_injectedExcel != null) AttachToExcel(_injectedExcel, false);
                else TryAttachToRunningExcel();
                return;
            }
            if (!ExcelWindowIsAlive())
            {
                InvalidateCurrentPopup();
                DetachFromExcel();
                SetStatus("Excel 已关闭，正在等待重新打开…");
                return;
            }
            try
            {
                if (!_excel.Ready)
                {
                    SetStatus("Excel 正在编辑或忙碌，工具会自动重试…");
                    return;
                }
                EnsureCommandButtons();
            }
            catch (COMException error)
            {
                if (IsDisconnected(error))
                {
                    InvalidateCurrentPopup();
                    DetachFromExcel();
                    SetStatus("Excel 连接已断开，正在重新连接…");
                }
                else SetStatus("Excel 正在忙碌，工具会自动重试…");
            }
        }

        private void TryAttachToRunningExcel()
        {
            Excel.Application app = null;
            try
            {
                object running = Marshal.GetActiveObject("Excel.Application");
                app = running as Excel.Application;
                if (app == null)
                {
                    SetStatus("正在等待 Excel 就绪…");
                    ReleaseComObject(running);
                    return;
                }
                try
                {
                    if (!app.Ready)
                    {
                        SetStatus("Excel 正在编辑或忙碌，工具会自动重试…");
                        ReleaseComObject(app);
                        return;
                    }
                }
                catch (COMException)
                {
                    SetStatus("Excel 正在忙碌，工具会自动重试…");
                    ReleaseComObject(app);
                    return;
                }
                AttachToExcel(app, true);
            }
            catch (COMException)
            {
                if (app != null && !object.ReferenceEquals(app, _excel)) ReleaseComObject(app);
                SetStatus("未检测到 Excel。打开 Excel 后会自动连接。");
            }
            catch (InvalidCastException)
            {
                if (app != null && !object.ReferenceEquals(app, _excel)) ReleaseComObject(app);
                SetStatus("正在等待可用的 Excel 实例…");
            }
        }

        private void AttachToExcel(Excel.Application app, bool ownsReference)
        {
            if (app == null || _disposed) return;
            _excel = app;
            _ownsExcelReference = ownsReference;
            try
            {
                _excelHwnd = _excel.Hwnd;
                GetWindowThreadProcessId(new IntPtr(_excelHwnd), out _excelProcessId);
                _excel.SheetSelectionChange += Excel_SheetSelectionChange;
                _selectionEventAttached = true;
                _excel.SheetBeforeRightClick += Excel_SheetBeforeRightClick;
                _rightClickEventAttached = true;
                _excel.SheetActivate += Excel_SheetActivate;
                _sheetActivateEventAttached = true;
                _excel.WindowDeactivate += Excel_WindowDeactivate;
                _windowDeactivateEventAttached = true;
                EnsureCommandButtons();
                if (HasAnyButton()) SetStatus("已连接 Excel；支持一个 Excel 实例。");
                else SetStatus("已连接 Excel，正在等待右键菜单就绪…");
            }
            catch (COMException ex)
            {
                DetachFromExcel();
                SetStatus("Excel 正在忙碌或尚未就绪，工具会自动重试。 (" + ex.ErrorCode.ToString("X8") + ")");
            }
            catch (InvalidCastException)
            {
                DetachFromExcel();
                SetStatus("无法连接此 Excel 实例，工具会自动重试。");
            }
        }

        private bool HasAnyButton()
        {
            return _cellButton != null || _listButton != null;
        }

        private void EnsureCommandButtons()
        {
            if (_excel == null || _disposed) return;
            try
            {
                if (!_staleTagsCleaned)
                {
                    CleanupStaleMenuTags("Cell");
                    CleanupStaleMenuTags("List Range Popup");
                    CleanupStaleMenuTags("Formula Bar");
                    _staleTagsCleaned = true;
                }
                if (_cellButton == null) _cellButton = AddMenuButton("Cell");
                if (_listButton == null) _listButton = AddMenuButton("List Range Popup");
                if (_editButton == null) _editButton = AddMenuButton("Formula Bar");
                if (HasAnyButton()) SetStatus("已连接 Excel；支持一个 Excel 实例。");
            }
            catch (COMException)
            {
                if (!HasAnyButton()) SetStatus("已连接 Excel，正在等待右键菜单就绪…");
            }
        }

        private Office.CommandBarButton AddMenuButton(string barName)
        {
            Office.CommandBars bars = null;
            Office.CommandBar bar = null;
            Office.CommandBarControls controls = null;
            Office.CommandBarControl control = null;
            try
            {
                bars = _excel.CommandBars;
                bar = bars[barName];
                if (bar == null) return null;
                controls = bar.Controls;
                control = controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                Office.CommandBarButton button = control as Office.CommandBarButton;
                if (button == null) return null;
                bool editText = string.Equals(barName, "Formula Bar", StringComparison.OrdinalIgnoreCase);
                button.Caption = editText ? "翻译所选文字（Ctrl+Alt+T）" : MenuCaption;
                button.Tag = editText ? EditMenuTag : string.Equals(barName, "Cell", StringComparison.OrdinalIgnoreCase) ? _menuTag : _tableMenuTag;
                button.TooltipText = editText ? "中英自动切换；只翻译编辑区选中的文字" : "中英文自动识别并翻译所选单元格";
                button.Style = Office.MsoButtonStyle.msoButtonCaption;
                button.BeginGroup = true;
                button.Visible = true;
                button.Enabled = true;
                if (editText) button.Click += EditMenuButton_Click;
                else button.Click += MenuButton_Click;
                control = null;
                return button;
            }
            finally
            {
                ReleaseComObject(control);
                ReleaseComObject(controls);
                ReleaseComObject(bar);
                ReleaseComObject(bars);
            }
        }

        private void CleanupStaleMenuTags(string barName)
        {
            Office.CommandBars bars = null;
            Office.CommandBar bar = null;
            Office.CommandBarControls controls = null;
            try
            {
                bars = _excel.CommandBars;
                bar = bars[barName];
                if (bar == null) return;
                controls = bar.Controls;
                for (int index = controls.Count; index >= 1; index--)
                {
                    Office.CommandBarControl control = null;
                    try
                    {
                        control = controls[index];
                        string tag = control.Tag as string;
                        if (!string.IsNullOrEmpty(tag) && tag.StartsWith(MenuPrefix, StringComparison.OrdinalIgnoreCase))
                            control.Delete(Type.Missing);
                    }
                    finally { ReleaseComObject(control); }
                }
            }
            catch (COMException) { }
            finally
            {
                ReleaseComObject(controls);
                ReleaseComObject(bar);
                ReleaseComObject(bars);
            }
        }

        private void Excel_SheetSelectionChange(object sheet, Excel.Range target)
        {
            Interlocked.Increment(ref _eventRevision);
            RunOnUi(delegate { InvalidateCurrentPopup(); });
        }

        private void Excel_SheetBeforeRightClick(object sheet, Excel.Range target, ref bool cancel)
        {
            Interlocked.Increment(ref _eventRevision);
            // Do not retain callback COM arguments. The click reads the current selection on the UI STA.
            RunOnUi(delegate { InvalidateCurrentPopup(); _popupAnchor = Cursor.Position; });
        }

        private void Excel_SheetActivate(object sheet)
        {
            Interlocked.Increment(ref _eventRevision);
            RunOnUi(delegate { InvalidateCurrentPopup(); });
        }

        private void Excel_WindowDeactivate(Excel.Workbook workbook, Excel.Window window)
        {
            Interlocked.Increment(ref _eventRevision);
            RunOnUi(delegate { InvalidateCurrentPopup(); });
        }

        private void MenuButton_Click(Office.CommandBarButton button, ref bool cancelDefault)
        {
            long eventRevision = Interlocked.Read(ref _eventRevision);
            RunOnUi(delegate
            {
                if (_excel != null && eventRevision == Interlocked.Read(ref _eventRevision))
                    TranslateContextCellAsync(eventRevision);
            });
        }

        private void EditMenuButton_Click(Office.CommandBarButton button, ref bool cancelDefault)
        {
            cancelDefault = true;
            long revision = Interlocked.Read(ref _eventRevision);
            EditSelectionBridge bridge = _editSelection;
            long inputRevision = bridge == null ? 0 : bridge.InputRevision;
            // Never read Excel.Selection here: while editing it is a cell range,
            // not the selected substring, and Excel may reject COM calls.
            RunOnUi(delegate
            {
                if (revision != Interlocked.Read(ref _eventRevision)) return;
                Action requested = NativeTextTranslationRequested;
                if (requested != null) requested();
                if (bridge != null && object.ReferenceEquals(bridge, _editSelection))
                    bridge.TranslateSelectionFromNativeMenu(inputRevision);
            });
        }

        private void RunOnUi(Action action)
        {
            if (_disposed) return;
            if (Thread.CurrentThread.ManagedThreadId == _uiThreadId) { action(); return; }
            Control dispatcher = _dispatcher;
            if (dispatcher == null || dispatcher.IsDisposed) return;
            try { dispatcher.BeginInvoke(new Action(delegate { if (!_disposed) action(); })); }
            catch (InvalidOperationException) { /* Shutdown raced with an Office callback. */ }
        }

        private async void TranslateContextCellAsync(long eventRevision)
        {
            long generation = BeginOperation();
            string source = null;
            string error = null;
            Point anchor = _popupAnchor;
            Excel.Workbook workbook = null;
            Excel.Range selection = null;
            Excel.Range selectionCells = null;
            Excel.Range firstCell = null;
            Excel.Range mergedRange = null;
            Excel.Range mergedCells = null;
            Excel.Range topLeft = null;
            Excel.Areas areas = null;
            object selectionObject = null;
            bool isMergedSelection = false;
            try
            {
                selectionObject = _excel.Selection;
                selection = selectionObject as Excel.Range;
                if (selection == null)
                {
                    error = "请先在工作表中选中一个单元格，再打开右键菜单。";
                }
                else
                {
                    areas = selection.Areas;
                    if (areas == null || areas.Count != 1)
                    {
                        error = "请选择单个单元格，或右键单个合并单元格。";
                    }
                    else
                    {
                        workbook = _excel.ActiveWorkbook;
                        if (workbook == null)
                        {
                            error = "当前工作簿已切换，请重新选择单元格后翻译。";
                        }
                        else
                        {
                            selectionCells = selection.Cells;
                            firstCell = selectionCells[1, 1] as Excel.Range;
                            if (firstCell == null) error = "无法读取所选单元格，请重新选择后重试。";
                            else
                            {
                                object merged = firstCell.MergeCells;
                                isMergedSelection = merged is bool && (bool)merged;
                                if (selection.CountLarge > 1 && !isMergedSelection)
                                    error = "请选择单个单元格，或右键单个合并单元格。";
                                else if (isMergedSelection)
                                {
                                    mergedRange = firstCell.MergeArea;
                                    if (mergedRange == null) error = "无法读取所选合并单元格，请重新选择后重试。";
                                    else
                                    {
                                        if (selection.CountLarge > 1)
                                        {
                                            string selectedAddress = selection.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                                            string mergedAddress = mergedRange.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                                            if (!string.Equals(selectedAddress, mergedAddress, StringComparison.OrdinalIgnoreCase))
                                                error = "请选择单个单元格，或右键单个合并单元格。";
                                        }
                                        mergedCells = mergedRange.Cells;
                                        topLeft = mergedCells[1, 1] as Excel.Range;
                                    }
                                }

                                if (error == null)
                                {
                                    Excel.Range valueCell = isMergedSelection ? topLeft : firstCell;
                                    if (valueCell == null) error = "无法读取所选单元格，请重新选择后重试。";
                                    else
                                    {
                                        object value = valueCell.Value2;
                                        if (value is ErrorWrapper) error = "这个单元格包含 Excel 错误值，无法翻译。";
                                        else if (!(value is string)) error = value == null ? "这个单元格是空的，没有可翻译的文字。" : "这个单元格不是文字，请选择包含文字的单元格。";
                                        else
                                        {
                                            source = (string)value;
                                            if (string.IsNullOrWhiteSpace(source)) error = "这个单元格是空的，没有可翻译的文字。";
                                            else if (source.Length > TranslationService.MaximumCharacters) error = "文字超过 2,000 字符，请先缩短单元格内容。";
                                            else if (TryGetAnchor(isMergedSelection ? mergedRange : firstCell, out anchor)) _popupAnchor = anchor;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (COMException)
            {
                error = "Excel 正在忙碌或单元格已经改变，请重新选择后重试。";
            }
            finally
            {
                ReleaseComObject(topLeft);
                ReleaseComObject(mergedCells);
                ReleaseComObject(mergedRange);
                ReleaseComObject(firstCell);
                ReleaseComObject(selectionCells);
                ReleaseComObject(areas);
                ReleaseComObject(selection);
                if (!object.ReferenceEquals(selectionObject, selection)) ReleaseComObject(selectionObject);
                ReleaseComObject(workbook);
            }

            _lastSourceText = source;
            if (_disposed || generation != _generation || eventRevision != Interlocked.Read(ref _eventRevision)) return;
            if (error != null)
            {
                if (generation == _generation) ShowPopup("无法翻译", error, false, anchor);
                return;
            }

            await TranslateSourceAsync(source, anchor, generation, eventRevision);
        }

        private async void TranslateSelectedText(string source, Point anchor)
        {
            long generation = BeginOperation();
            long revision = Interlocked.Read(ref _eventRevision);
            _lastSourceText = source;
            _popupAnchor = anchor;
            await TranslateSourceAsync(source, anchor, generation, revision);
        }

        private async Task TranslateSourceAsync(string source, Point anchor, long generation, long eventRevision)
        {
            var cancellation = new CancellationTokenSource();
            _activeCancellation = cancellation;
            EnsureDismissHooks();
            if (_showProgressPopup && generation == _generation) ShowPopup("正在翻译…", "正在通过所选翻译服务翻译文字。", true, anchor);
            try
            {
                TranslationResult result = await _translate(source, cancellation.Token);
                if (_disposed || cancellation.IsCancellationRequested || generation != _generation || eventRevision != Interlocked.Read(ref _eventRevision) || result == null) return;
                string heading = result.Direction;
                if (!string.IsNullOrEmpty(result.Provider)) heading += " · " + result.Provider.Replace("Google Cloud Translation", "Google");
                ShowPopup(heading, result.Text, false, anchor);
                SetStatus("翻译完成；已连接 Excel。");
            }
            catch (OperationCanceledException)
            {
                // A selection change, deactivation, or a newer request superseded this result.
            }
            catch (Exception ex)
            {
                if (!_disposed && !cancellation.IsCancellationRequested && generation == _generation && eventRevision == Interlocked.Read(ref _eventRevision))
                {
                    ShowPopup("翻译未完成", string.IsNullOrWhiteSpace(ex.Message) ? "请检查网络后重试。" : ex.Message, false, anchor);
                    SetStatus("翻译未完成；请检查网络后重试。");
                }
            }
            finally
            {
                if (object.ReferenceEquals(_activeCancellation, cancellation)) _activeCancellation = null;
                cancellation.Dispose();
            }
        }

        private long BeginOperation()
        {
            CancelActiveTranslation();
            long generation = Interlocked.Increment(ref _generation);
            HidePopup();
            _lastSourceText = null;
            return generation;
        }

        private void InvalidateCurrentPopup()
        {
            Interlocked.Increment(ref _generation);
            CancelActiveTranslation();
            HidePopup();
        }

        private void CancelActiveTranslation()
        {
            CancellationTokenSource cancellation = _activeCancellation;
            _activeCancellation = null;
            if (cancellation != null)
            {
                try { cancellation.Cancel(); } catch (ObjectDisposedException) { }
            }
        }

        private void ShowPopup(string heading, string message, bool loading, Point anchor)
        {
            HidePopup();
            _popup = new TranslationPopupForm(heading, message, loading);
            _popup.Dismissed += Popup_Dismissed;
            Point location = ClampPopupLocation(_popup, anchor);
            _popup.Location = location;
            _popup.Show();
            EnsureDismissHooks();
        }

        private void EnsureDismissHooks()
        {
            // Deterministic COM tests control focus and must not hook real user input.
            if (_foregroundCheck != null) return;
            if (_dismissHooks != null) return;
            _dismissHooks = new GlobalDismissHooks(DismissFromKeyboard, DismissFromMouse);
            _dismissHooks.Install();
        }

        private void Popup_Dismissed(object sender, EventArgs e)
        {
            InvalidateCurrentPopup();
        }

        private void DismissFromKeyboard()
        {
            if (_disposed || (_popup == null && _activeCancellation == null)) return;
            if (!ForegroundBelongsToExcel()) return;
            InvalidateCurrentPopup();
        }

        private void DismissFromMouse(Point pointer)
        {
            if (_disposed || (_popup == null && _activeCancellation == null)) return;
            if (_popup != null && _popup.Visible && _popup.Bounds.Contains(pointer)) return;
            if (!ForegroundBelongsToExcel()) return;
            InvalidateCurrentPopup();
        }

        private void HidePopup()
        {
            if (_dismissHooks != null)
            {
                _dismissHooks.Dispose();
                _dismissHooks = null;
            }
            TranslationPopupForm popup = _popup;
            _popup = null;
            if (popup != null)
            {
                try { popup.Hide(); } catch (ObjectDisposedException) { }
                popup.Dispose();
            }
        }

        private void ForegroundTimer_Tick(object sender, EventArgs e)
        {
            if ((_popup != null && _popup.Visible || _activeCancellation != null) && !ForegroundBelongsToExcel())
            {
                InvalidateCurrentPopup();
            }
        }

        private bool ForegroundBelongsToExcel()
        {
            if (_foregroundCheck != null) return _foregroundCheck();
            if (_excelProcessId == 0) return false;
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;
            uint processId;
            GetWindowThreadProcessId(foreground, out processId);
            return processId == _excelProcessId;
        }

        private bool ExcelWindowIsAlive()
        {
            if (_excelHwnd == 0 || _excelProcessId == 0) return false;
            IntPtr window = new IntPtr(_excelHwnd);
            if (!IsWindow(window)) return false;
            uint processId;
            uint threadId = GetWindowThreadProcessId(window, out processId);
            return threadId != 0 && processId == _excelProcessId;
        }

        private static bool IsDisconnected(COMException error)
        {
            int code = error.ErrorCode;
            return code == unchecked((int)0x80010108) || // RPC_E_DISCONNECTED
                   code == unchecked((int)0x80010007) || // RPC_E_SERVER_DIED
                   code == unchecked((int)0x800401FD);   // CO_E_OBJNOTCONNECTED
        }

        private bool TryGetAnchor(Excel.Range range, out Point anchor)
        {
            anchor = Cursor.Position;
            Excel.Window window = null;
            try
            {
                window = _excel.ActiveWindow;
                if (window == null) return false;
                double zoom = Convert.ToDouble(window.Zoom, System.Globalization.CultureInfo.InvariantCulture) / 100.0;
                double scale = zoom * WindowDpi(new IntPtr(window.Hwnd)) / 72.0;
                int right = window.PointsToScreenPixelsX(0) + (int)Math.Round((range.Left + range.Width) * scale, MidpointRounding.AwayFromZero);
                int top = window.PointsToScreenPixelsY(0) + (int)Math.Round(range.Top * scale, MidpointRounding.AwayFromZero);
                anchor = new Point(right, top);
                return true;
            }
            catch (COMException) { return false; }
            finally { ReleaseComObject(window); }
        }

        internal static double WindowDpi(IntPtr hwnd)
        {
            try { uint dpi = GetDpiForWindow(hwnd); if (dpi > 0) return dpi; }
            catch (EntryPointNotFoundException) { }
            using (Graphics graphics = Graphics.FromHwnd(hwnd)) return graphics.DpiX;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        private static Point ClampPopupLocation(TranslationPopupForm popup, Point anchor)
        {
            Rectangle working = Screen.FromPoint(anchor).WorkingArea;
            int x = anchor.X + 8;
            if (x + popup.Width > working.Right) x = anchor.X - popup.Width - 8;
            int y = anchor.Y;
            if (y + popup.Height > working.Bottom) y = working.Bottom - popup.Height - 8;
            if (y < working.Top + 4) y = working.Top + 4;
            if (x < working.Left + 4) x = working.Left + 4;
            if (x + popup.Width > working.Right - 4) x = working.Right - popup.Width - 4;
            return new Point(x, y);
        }

        private void SetStatus(string text)
        {
            if (_disposed || string.Equals(_statusText, text, StringComparison.Ordinal)) return;
            _statusText = text;
            EventHandler handler = StatusChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void DetachFromExcel()
        {
            Excel.Application app = _excel;
            if (app == null) return;
            if (_selectionEventAttached) { try { app.SheetSelectionChange -= Excel_SheetSelectionChange; } catch (COMException) { } }
            if (_rightClickEventAttached) { try { app.SheetBeforeRightClick -= Excel_SheetBeforeRightClick; } catch (COMException) { } }
            if (_sheetActivateEventAttached) { try { app.SheetActivate -= Excel_SheetActivate; } catch (COMException) { } }
            if (_windowDeactivateEventAttached) { try { app.WindowDeactivate -= Excel_WindowDeactivate; } catch (COMException) { } }
            _selectionEventAttached = _rightClickEventAttached = _sheetActivateEventAttached = _windowDeactivateEventAttached = false;
            RemoveButton(ref _cellButton);
            RemoveButton(ref _listButton);
            RemoveButton(ref _editButton);
            CleanupStaleMenuTags("Cell");
            CleanupStaleMenuTags("List Range Popup");
            CleanupStaleMenuTags("Formula Bar");
            _excel = null;
            uint oldPid = _excelProcessId;
            _excelHwnd = 0;
            _excelProcessId = 0;
            _staleTagsCleaned = false;
            if (_ownsExcelReference) ReleaseComObject(app);
            _ownsExcelReference = false;
            GC.KeepAlive(oldPid);
        }

        private void RemoveButton(ref Office.CommandBarButton button)
        {
            Office.CommandBarButton owned = button;
            button = null;
            if (owned == null) return;
            try { owned.Click -= MenuButton_Click; } catch (COMException) { }
            try { owned.Click -= EditMenuButton_Click; } catch (COMException) { }
            try { owned.Delete(Type.Missing); } catch (COMException) { }
            ReleaseComObject(owned);
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.ReleaseComObject(value); } catch (InvalidComObjectException) { }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Interlocked.Increment(ref _generation);
            CancelActiveTranslation();
            if (_attachTimer != null) { _attachTimer.Stop(); _attachTimer.Dispose(); _attachTimer = null; }
            if (_foregroundTimer != null) { _foregroundTimer.Stop(); _foregroundTimer.Dispose(); _foregroundTimer = null; }
            HidePopup();
            if (_editSelection != null) { _editSelection.Dispose(); _editSelection = null; }
            DetachFromExcel();
            if (_dispatcher != null) { _dispatcher.Dispose(); _dispatcher = null; }
            StatusChanged = null;
            NativeTextTranslationRequested = null;
        }

        // Test harnesses and the production connection manager share this event path.
        internal void ProcessSelectionChangeForTesting()
        {
            Excel_SheetSelectionChange(null, null);
        }


        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr window);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    }

    internal sealed class TranslationPopupForm : Form
    {
        private readonly Label _heading;
        private readonly Label _body;
        internal event EventHandler Dismissed;

        internal TranslationPopupForm(string heading, string message, bool loading)
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = Color.White;
            Padding = new Padding(0);
            Font = new Font("Microsoft YaHei UI", 9F);
            Width = 360;
            int textHeight = TextRenderer.MeasureText(message ?? string.Empty, Font,
                new Size(306, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix |
                TextFormatFlags.TextBoxControl).Height;
            int bodyHeight = Math.Max(54, Math.Min(300, textHeight + 12));
            Height = 74 + bodyHeight;
            _heading = new Label
            {
                Text = heading,
                Location = new Point(18, 15),
                Size = new Size(290, 25),
                AutoEllipsis = true,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                ForeColor = loading ? Color.FromArgb(45, 105, 185) : Color.FromArgb(39, 55, 79)
            };
            var close = new Button
            {
                Text = "×",
                Location = new Point(326, 10),
                Size = new Size(24, 24),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.FromArgb(121, 132, 149),
                BackColor = Color.White,
                TabStop = false
            };
            close.FlatAppearance.BorderSize = 0;
            close.Click += delegate { EventHandler handler = Dismissed; if (handler != null) handler(this, EventArgs.Empty); };
            var separator = new Panel { Location = new Point(16, 46), Size = new Size(328, 1), BackColor = Color.FromArgb(230, 234, 241) };
            var bodyPanel = new Panel
            {
                Location = new Point(14, 57),
                Size = new Size(332, bodyHeight),
                AutoScroll = true,
                AutoScrollMinSize = new Size(0, textHeight + 12),
                BackColor = Color.White
            };
            _body = new Label
            {
                Text = message,
                Location = new Point(4, 4),
                Size = new Size(306, textHeight + 4),
                ForeColor = Color.FromArgb(59, 71, 91),
                Font = new Font("Microsoft YaHei UI", 9F),
                AutoSize = false,
                UseMnemonic = false
            };
            bodyPanel.Controls.Add(_body);
            Controls.Add(_heading); Controls.Add(close); Controls.Add(separator); Controls.Add(bodyPanel);
            using (var path = RoundedRectangle(new Rectangle(0, 0, Width, Height), 9))
                Region = new Region(path);
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams parameters = base.CreateParams;
                parameters.ExStyle |= 0x08000000 | 0x00000080; // WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW
                return parameters;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using (var pen = new Pen(Color.FromArgb(220, 226, 236)))
            using (var path = RoundedRectangle(new Rectangle(0, 0, Width - 1, Height - 1), 9))
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(pen, path);
            }
            base.OnPaint(e);
        }

        private static System.Drawing.Drawing2D.GraphicsPath RoundedRectangle(Rectangle bounds, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            int diameter = radius * 2;
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    internal sealed class GlobalDismissHooks : IDisposable
    {
        private const int WhKeyboardLl = 13;
        private const int WhMouseLl = 14;
        private const int WmKeyDown = 0x0100;
        private const int WmSysKeyDown = 0x0104;
        private const int WmMouseWheel = 0x020A;
        private const int WmMouseHWheel = 0x020E;
        private const int VkEscape = 0x1B;
        private const int VkPageUp = 0x21;
        private const int VkPageDown = 0x22;
        private const int VkUp = 0x26;
        private const int VkDown = 0x28;
        private readonly Action _keyboardDismiss;
        private readonly Action<Point> _mouseDismiss;
        private readonly LowLevelKeyboardProc _keyboardProc;
        private readonly LowLevelMouseProc _mouseProc;
        private IntPtr _keyboardHook;
        private IntPtr _mouseHook;

        internal GlobalDismissHooks(Action keyboardDismiss, Action<Point> mouseDismiss)
        {
            _keyboardDismiss = keyboardDismiss;
            _mouseDismiss = mouseDismiss;
            _keyboardProc = KeyboardCallback;
            _mouseProc = MouseCallback;
        }

        internal void Install()
        {
            if (_keyboardHook != IntPtr.Zero || _mouseHook != IntPtr.Zero) return;
            _keyboardHook = SetWindowsHookEx(WhKeyboardLl, _keyboardProc, GetModuleHandle(null), 0);
            _mouseHook = SetWindowsHookEx(WhMouseLl, _mouseProc, GetModuleHandle(null), 0);
        }

        private IntPtr KeyboardCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && (wParam.ToInt32() == WmKeyDown || wParam.ToInt32() == WmSysKeyDown))
            {
                var key = (KeyboardHookData)Marshal.PtrToStructure(lParam, typeof(KeyboardHookData));
                if (key.VirtualKey == VkEscape || key.VirtualKey == VkPageUp || key.VirtualKey == VkPageDown || key.VirtualKey == VkUp || key.VirtualKey == VkDown)
                    _keyboardDismiss();
            }
            return CallNextHookEx(_keyboardHook, code, wParam, lParam);
        }

        private IntPtr MouseCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && (wParam.ToInt32() == WmMouseWheel || wParam.ToInt32() == WmMouseHWheel))
            {
                var mouse = (MouseHookData)Marshal.PtrToStructure(lParam, typeof(MouseHookData));
                _mouseDismiss(mouse.Point);
            }
            return CallNextHookEx(_mouseHook, code, wParam, lParam);
        }

        public void Dispose()
        {
            if (_keyboardHook != IntPtr.Zero) { UnhookWindowsHookEx(_keyboardHook); _keyboardHook = IntPtr.Zero; }
            if (_mouseHook != IntPtr.Zero) { UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardHookData
        {
            internal uint VirtualKey;
            internal uint ScanCode;
            internal uint Flags;
            internal uint Time;
            internal IntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MouseHookData
        {
            internal Point Point;
            internal uint MouseData;
            internal uint Flags;
            internal uint Time;
            internal IntPtr ExtraInfo;
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
    }

    internal static class AppIcon
    {
        internal static Icon Create()
        {
            using (var bitmap = new Bitmap(32, 32))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var brush = new SolidBrush(Color.FromArgb(48, 102, 177)))
            using (var font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.FillRoundedRectangle(brush, new Rectangle(1, 1, 30, 30), 6);
                using (var textBrush = new SolidBrush(Color.White))
                    graphics.DrawString("译", font, textBrush, new PointF(5, 6));
                IntPtr handle = bitmap.GetHicon();
                try { return (Icon)Icon.FromHandle(handle).Clone(); }
                finally { DestroyIcon(handle); }
            }
        }

        private static void FillRoundedRectangle(this Graphics graphics, Brush brush, Rectangle bounds, int radius)
        {
            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                int d = radius * 2;
                path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                graphics.FillPath(brush, path);
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr icon);
    }
}
