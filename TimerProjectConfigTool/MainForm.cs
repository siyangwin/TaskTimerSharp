using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace TimerProjectConfigTool
{
    public interface ILocalizableForm
    {
        void ApplyLanguage();
    }

    public class MainForm : Form, ILocalizableForm
    {
        private string _rootDir;

        // 顶栏
        private Label _lblRootDir;
        private TextBox _txtRootDir;
        private Button _btnBrowse;
        private Button _btnRefresh;
        private Label _lblLang;
        private ComboBox _cboLang;
        private Label _lblServiceStatusTitle;
        private Label _lblServiceStatus;
        private Button _btnRefreshStatus;
        private Button _btnRestart;

        // 选项卡
        private TabControl _tabs;
        private TabPage _tabJobs;
        private TabPage _tabLogs;
        private TabPage _tabMail;

        // 任务管理
        private DataGridView _grid;
        private Button _btnNewJob;
        private Button _btnEdit;
        private Button _btnEnable;
        private Button _btnDisable;
        private Button _btnOpenFolder;

        // 日志浏览
        private Button _btnRefreshLogs;
        private TreeView _tree;
        private TextBox _txtPreview;

        // 邮件设置
        private Label _lblMailNotice;
        private Label _lblMailHost;
        private Label _lblPort;
        private Label _lblMailAddress;
        private Label _lblMailDisplayName;
        private Label _lblMailPassword;
        private Label _lblSsl;
        private Label _lblProgramName;
        private Label _lblSendTo;
        private Label _lblInfoMessage;
        private TextBox _txtMailHost;
        private TextBox _txtPort;
        private TextBox _txtMailAddress;
        private TextBox _txtMailDisplayName;
        private TextBox _txtMailPassword;
        private CheckBox _chkShowPassword;
        private ComboBox _cboSsl;
        private TextBox _txtProgramName;
        private TextBox _txtSendTo;
        private ComboBox _cboInfoMessage;
        private Button _btnSaveMail;
        private Button _btnTestSend;

        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _statusLabel;

        private bool _loadingLangCombo;

        public MainForm()
        {
            string savedRoot = SettingsStore.LoadRootDir();
            _rootDir = (!string.IsNullOrEmpty(savedRoot) && Directory.Exists(savedRoot))
                ? savedRoot
                : AppDomain.CurrentDomain.BaseDirectory;
            BuildUi();
            ApplyLanguage();
        }

        private void BuildUi()
        {
            Text = "TimerProjectConfigTool";
            Size = new Size(1100, 720);
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;

            // ---------- 顶栏 ----------
            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 8,
                AutoSize = true,
                Padding = new Padding(8, 6, 8, 6)
            };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            top.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _lblRootDir = new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 3, 3) };
            _txtRootDir = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 5, 3, 3) };
            _txtRootDir.Text = _rootDir;
            _txtRootDir.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    ApplyRootDir(_txtRootDir.Text);
                }
            };
            _btnBrowse = new Button { AutoSize = true, Margin = new Padding(3, 4, 3, 3) };
            _btnBrowse.Click += BtnBrowse_Click;
            _btnRefresh = new Button { AutoSize = true, Margin = new Padding(3, 4, 3, 3) };
            _btnRefresh.Click += (s, e) => RefreshAll();

            _lblLang = new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 3, 3) };
            _cboLang = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130, Margin = new Padding(3, 5, 3, 3) };
            _cboLang.SelectedIndexChanged += CboLang_SelectedIndexChanged;

            _lblServiceStatusTitle = new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(16, 8, 3, 3) };
            _lblServiceStatus = new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 3, 3), Font = new Font(Font.FontFamily, Font.Size, FontStyle.Bold) };
            _btnRefreshStatus = new Button { AutoSize = true, Margin = new Padding(3, 4, 3, 3) };
            _btnRefreshStatus.Click += (s, e) => RefreshServiceStatus();
            _btnRestart = new Button { AutoSize = true, Margin = new Padding(3, 4, 3, 3) };
            _btnRestart.Click += async (s, e) => await RestartServiceAsync();

            top.Controls.Add(_lblRootDir, 0, 0);
            top.Controls.Add(_txtRootDir, 1, 0);
            top.Controls.Add(_btnBrowse, 2, 0);
            top.Controls.Add(_btnRefresh, 3, 0);
            top.Controls.Add(_lblLang, 4, 0);
            top.Controls.Add(_cboLang, 5, 0);
            top.Controls.Add(_lblServiceStatusTitle, 6, 0);
            top.Controls.Add(_lblServiceStatus, 7, 0);
            top.Controls.Add(_btnRefreshStatus, 6, 1);
            top.Controls.Add(_btnRestart, 7, 1);

            // ---------- 状态栏 ----------
            _statusStrip = new StatusStrip();
            _statusLabel = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            _statusStrip.Items.Add(_statusLabel);

            // ---------- 选项卡 ----------
            _tabs = new TabControl { Dock = DockStyle.Fill };
            _tabJobs = new TabPage();
            _tabLogs = new TabPage();
            _tabMail = new TabPage();
            _tabs.TabPages.Add(_tabJobs);
            _tabs.TabPages.Add(_tabLogs);
            _tabs.TabPages.Add(_tabMail);

            BuildJobsTab();
            BuildLogsTab();
            BuildMailTab();

            Controls.Add(_tabs);
            Controls.Add(top);
            Controls.Add(_statusStrip);
        }

        // ================= 任务管理 =================
        private void BuildJobsTab()
        {
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.Fixed3D
            };
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "NameCN", Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "NameEN", Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", Width = 110 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", Width = 70 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Folder", Width = 70 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Setup", Width = 90 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Sequence", Width = 100 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Issue", Width = 150 });
            _grid.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex >= 0) EditSelectedJob();
            };
            _grid.SelectionChanged += (s, e) => UpdateJobButtons();

            _btnNewJob = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnNewJob.Click += BtnNewJob_Click;
            _btnEdit = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnEdit.Click += (s, e) => EditSelectedJob();
            _btnEnable = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnEnable.Click += (s, e) => SetSelectedJobEnabled(true);
            _btnDisable = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnDisable.Click += (s, e) => SetSelectedJobEnabled(false);
            _btnOpenFolder = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnOpenFolder.Click += BtnOpenFolder_Click;

            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(4)
            };
            bar.Controls.Add(_btnNewJob);
            bar.Controls.Add(_btnEdit);
            bar.Controls.Add(_btnEnable);
            bar.Controls.Add(_btnDisable);
            bar.Controls.Add(_btnOpenFolder);

            _tabJobs.Controls.Add(_grid);
            _tabJobs.Controls.Add(bar);
        }

        private JobOverviewRow SelectedRow()
        {
            if (_grid.CurrentRow == null) return null;
            return _grid.CurrentRow.Tag as JobOverviewRow;
        }

        private void UpdateJobButtons()
        {
            var row = SelectedRow();
            bool has = row != null;
            _btnEdit.Enabled = has;
            _btnEnable.Enabled = has && row.InRegistry && !row.Enabled;
            _btnDisable.Enabled = has && row.InRegistry && row.Enabled;
            _btnOpenFolder.Enabled = has && row.FolderExists;
        }

        private void RefreshJobs()
        {
            _grid.Rows.Clear();
            try
            {
                foreach (var r in ConfigModel.LoadOverview(_rootDir))
                {
                    int idx = _grid.Rows.Add(
                        r.NameCN,
                        r.NameEN,
                        TypeText(r.Type),
                        r.InRegistry ? (r.Enabled ? Lang.Get("common.enabled") : Lang.Get("common.disabled")) : "-",
                        r.FolderExists ? Lang.Get("common.exists") : Lang.Get("common.missing"),
                        r.SetupExists ? Lang.Get("common.exists") : Lang.Get("common.missing"),
                        r.SequenceExists ? Lang.Get("common.exists") : Lang.Get("common.missing"),
                        r.IssueKey.Length > 0 ? Lang.Get(r.IssueKey) : "");
                    _grid.Rows[idx].Tag = r;
                    if (!r.SetupExists) _grid.Rows[idx].Cells["Setup"].Style.ForeColor = Color.Firebrick;
                    if (r.IssueKey.Length > 0)
                    {
                        _grid.Rows[idx].Cells["Issue"].Style.ForeColor = Color.Firebrick;
                        _grid.Rows[idx].Cells["Issue"].ToolTipText = Lang.Get(r.IssueKey + "Tip");
                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus(Lang.Get("main.loadConfigFailed", ex.Message));
            }
            UpdateJobButtons();
        }

        private static string TypeText(string type)
        {
            if (type == "0") return Lang.Get("main.typeApi");
            if (type == "1") return Lang.Get("main.typeSequence");
            if (type == "2") return Lang.Get("main.typeSp");
            return type;
        }

        private void BtnNewJob_Click(object sender, EventArgs e)
        {
            using (var f = new NewJobForm(_rootDir))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    RefreshJobs();
                    SelectJobByName(f.CreatedNameEN);
                    SetStatus(Lang.Get("main.createdDisabled"));
                    EditSelectedJob();
                }
            }
        }

        private void SelectJobByName(string nameEN)
        {
            foreach (DataGridViewRow row in _grid.Rows)
            {
                var r = row.Tag as JobOverviewRow;
                if (r != null && string.Equals(r.NameEN, nameEN, StringComparison.OrdinalIgnoreCase))
                {
                    _grid.ClearSelection();
                    row.Selected = true;
                    _grid.CurrentCell = row.Cells[0];
                    return;
                }
            }
        }

        private void EditSelectedJob()
        {
            var row = SelectedRow();
            if (row == null)
            {
                MessageBox.Show(this, Lang.Get("main.selectJobFirst"), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            using (var f = new JobEditForm(_rootDir, row.NameEN))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    RefreshJobs();
                    SelectJobByName(row.NameEN);
                    SetStatus(Lang.Get("main.savedNoRestart"));
                }
            }
        }

        private void SetSelectedJobEnabled(bool enabled)
        {
            var row = SelectedRow();
            if (row == null || !row.InRegistry) return;
            try
            {
                var registry = ConfigModel.LoadRegistry(_rootDir);
                foreach (var entry in registry)
                {
                    if (string.Equals(entry.NameEN, row.NameEN, StringComparison.OrdinalIgnoreCase))
                    {
                        entry.Enabled = enabled;
                        break;
                    }
                }
                ConfigModel.SaveRegistry(_rootDir, registry);
                RefreshJobs();
                SelectJobByName(row.NameEN);
                SetStatus(Lang.Get("main.savedNoRestart"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("main.saveFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            var row = SelectedRow();
            if (row == null || !row.FolderExists) return;
            try
            {
                Process.Start(new ProcessStartInfo(ConfigModel.JobDir(_rootDir, row.NameEN)) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("main.openFolderFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================= 日志浏览 =================
        private void BuildLogsTab()
        {
            _btnRefreshLogs = new Button { AutoSize = true, Margin = new Padding(4) };
            _btnRefreshLogs.Click += (s, e) => RefreshLogTree();
            var logTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            logTop.Controls.Add(_btnRefreshLogs);

            _tree = new TreeView { Dock = DockStyle.Fill };
            _tree.AfterSelect += (s, e) =>
            {
                string path = e.Node.Tag as string;
                if (path != null && File.Exists(path)) ShowPreview(path);
            };
            _tree.NodeMouseDoubleClick += (s, e) =>
            {
                string path = e.Node.Tag as string;
                if (path != null && File.Exists(path)) OpenWithShell(path);
            };

            _txtPreview = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 9f)
            };

            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 320 };
            split.Panel1.Controls.Add(_tree);
            split.Panel2.Controls.Add(_txtPreview);

            _tabLogs.Controls.Add(split);
            _tabLogs.Controls.Add(logTop);
        }

        private void RefreshLogTree()
        {
            _tree.Nodes.Clear();
            var runRoot = _tree.Nodes.Add(Lang.Get("logs.rootRunning"));
            var hisRoot = _tree.Nodes.Add(Lang.Get("logs.rootHistory"));
            try
            {
                ScanDirInto(runRoot, ConfigModel.LogDir(_rootDir), "History");
            }
            catch
            {
            }
            try
            {
                ScanDirInto(hisRoot, ConfigModel.HistoryDir(_rootDir), null);
            }
            catch
            {
            }
            runRoot.Expand();
            hisRoot.Expand();
        }

        private static void ScanDirInto(TreeNode parent, string dir, string skipDirName)
        {
            if (!Directory.Exists(dir)) return;
            var dirs = new List<string>(Directory.GetDirectories(dir));
            dirs.Sort((a, b) => string.Compare(b, a, StringComparison.OrdinalIgnoreCase));
            foreach (string d in dirs)
            {
                string name = Path.GetFileName(d);
                if (skipDirName != null && string.Equals(name, skipDirName, StringComparison.OrdinalIgnoreCase)) continue;
                var node = parent.Nodes.Add(name);
                ScanDirInto(node, d, skipDirName);
            }
            var files = new List<string>(Directory.GetFiles(dir));
            files.Sort((a, b) => string.Compare(b, a, StringComparison.OrdinalIgnoreCase));
            foreach (string f in files)
            {
                var node = parent.Nodes.Add(Path.GetFileName(f));
                node.Tag = f;
            }
        }

        private void ShowPreview(string path)
        {
            try
            {
                const int maxBytes = 200 * 1024;
                bool truncated = false;
                string text;
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    if (fs.Length > maxBytes)
                    {
                        fs.Seek(-maxBytes, SeekOrigin.End);
                        truncated = true;
                    }
                    using (var sr = new StreamReader(fs, Encoding.UTF8))
                    {
                        text = sr.ReadToEnd();
                    }
                }
                _txtPreview.Text = (truncated ? Lang.Get("logs.previewTooLarge") + Environment.NewLine : "") + text;
            }
            catch (Exception ex)
            {
                _txtPreview.Text = Lang.Get("logs.readFailed", ex.Message);
            }
        }

        private void OpenWithShell(string path)
        {
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("logs.openFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================= 邮件设置 =================
        private void BuildMailTab()
        {
            _lblMailNotice = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 34,
                BackColor = Color.FromArgb(255, 243, 205),
                ForeColor = Color.FromArgb(120, 80, 0),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoScroll = true,
                Padding = new Padding(10)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            _lblMailHost = NewFieldLabel();
            _txtMailHost = NewTextBox();
            _lblPort = NewFieldLabel();
            _txtPort = NewTextBox();
            _lblMailAddress = NewFieldLabel();
            _txtMailAddress = NewTextBox();
            _lblMailDisplayName = NewFieldLabel();
            _txtMailDisplayName = NewTextBox();
            _lblMailPassword = NewFieldLabel();
            _txtMailPassword = NewTextBox();
            _txtMailPassword.UseSystemPasswordChar = true;
            _chkShowPassword = new CheckBox { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            _chkShowPassword.CheckedChanged += (s, e) => _txtMailPassword.UseSystemPasswordChar = !_chkShowPassword.Checked;
            _lblSsl = NewFieldLabel();
            _cboSsl = NewBoolCombo();
            _lblProgramName = NewFieldLabel();
            _txtProgramName = NewTextBox();
            _lblSendTo = NewFieldLabel();
            _txtSendTo = NewTextBox();
            _lblInfoMessage = NewFieldLabel();
            _cboInfoMessage = NewBoolCombo();

            _btnSaveMail = new Button { AutoSize = true, Margin = new Padding(3, 12, 8, 3) };
            _btnSaveMail.Click += BtnSaveMail_Click;
            _btnTestSend = new Button { AutoSize = true, Margin = new Padding(3, 12, 3, 3) };
            _btnTestSend.Click += BtnTestSend_Click;

            var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0) };
            buttons.Controls.Add(_btnSaveMail);
            buttons.Controls.Add(_btnTestSend);

            int row = 0;
            AddMailRow(grid, row++, _lblMailHost, _txtMailHost);
            AddMailRow(grid, row++, _lblPort, _txtPort);
            AddMailRow(grid, row++, _lblMailAddress, _txtMailAddress);
            AddMailRow(grid, row++, _lblMailDisplayName, _txtMailDisplayName);
            AddMailRow(grid, row++, _lblMailPassword, _txtMailPassword);
            grid.Controls.Add(_chkShowPassword, 1, row++);
            AddMailRow(grid, row++, _lblSsl, _cboSsl);
            AddMailRow(grid, row++, _lblProgramName, _txtProgramName);
            AddMailRow(grid, row++, _lblSendTo, _txtSendTo);
            AddMailRow(grid, row++, _lblInfoMessage, _cboInfoMessage);
            grid.Controls.Add(buttons, 1, row);

            _tabMail.Controls.Add(grid);
            _tabMail.Controls.Add(_lblMailNotice);
        }

        private static Label NewFieldLabel()
        {
            return new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 9, 10, 3) };
        }

        private static TextBox NewTextBox()
        {
            return new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3), MaximumSize = new Size(420, 0) };
        }

        private static ComboBox NewBoolCombo()
        {
            var cbo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, Margin = new Padding(3, 6, 3, 3) };
            cbo.Items.Add("true");
            cbo.Items.Add("false");
            cbo.SelectedIndex = 1;
            return cbo;
        }

        private static void AddMailRow(TableLayoutPanel grid, int row, Label label, Control control)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            grid.Controls.Add(label, 0, row);
            grid.Controls.Add(control, 1, row);
        }

        private void LoadMailSettings()
        {
            string path = ConfigModel.ServiceConfigPath(_rootDir);
            if (!File.Exists(path))
            {
                SetStatus(Lang.Get("mail.configNotFound"));
                return;
            }
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(path);
                _txtMailHost.Text = ReadAppSetting(doc, "MailHost");
                _txtPort.Text = ReadAppSetting(doc, "Port");
                _txtMailAddress.Text = ReadAppSetting(doc, "MailAddress");
                _txtMailDisplayName.Text = ReadAppSetting(doc, "MailDisplayName");
                _txtMailPassword.Text = ReadAppSetting(doc, "MailPassWord");
                SelectBoolCombo(_cboSsl, ReadAppSetting(doc, "SSL"));
                _txtProgramName.Text = ReadAppSetting(doc, "ProgramName");
                _txtSendTo.Text = ReadAppSetting(doc, "SendTo");
                SelectBoolCombo(_cboInfoMessage, ReadAppSetting(doc, "InfoMessage"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("mail.loadFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string ReadAppSetting(XmlDocument doc, string key)
        {
            XmlNode n = doc.SelectSingleNode("configuration/appSettings/add[@key='" + key + "']");
            return n != null && n.Attributes["value"] != null ? n.Attributes["value"].Value : "";
        }

        private static void SelectBoolCombo(ComboBox cbo, string value)
        {
            cbo.SelectedIndex = string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
        }

        private async void BtnSaveMail_Click(object sender, EventArgs e)
        {
            string path = ConfigModel.ServiceConfigPath(_rootDir);
            if (!File.Exists(path))
            {
                MessageBox.Show(this, Lang.Get("mail.configNotFound"), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int port;
            if (!int.TryParse(_txtPort.Text.Trim(), out port) || port <= 0)
            {
                MessageBox.Show(this, Lang.Get("mail.portInvalid"), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(path);
                WriteAppSetting(doc, "MailHost", _txtMailHost.Text.Trim());
                WriteAppSetting(doc, "Port", _txtPort.Text.Trim());
                WriteAppSetting(doc, "MailAddress", _txtMailAddress.Text.Trim());
                WriteAppSetting(doc, "MailDisplayName", _txtMailDisplayName.Text.Trim());
                WriteAppSetting(doc, "MailPassWord", _txtMailPassword.Text);
                WriteAppSetting(doc, "SSL", _cboSsl.SelectedIndex == 0 ? "true" : "false");
                WriteAppSetting(doc, "ProgramName", _txtProgramName.Text.Trim());
                WriteAppSetting(doc, "SendTo", _txtSendTo.Text.Trim());
                WriteAppSetting(doc, "InfoMessage", _cboInfoMessage.SelectedIndex == 0 ? "true" : "false");

                string tmp = path + ".tmp_" + Guid.NewGuid().ToString("N");
                try
                {
                    doc.Save(tmp);
                    ConfigModel.AtomicReplace(tmp, path);
                }
                finally
                {
                    try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("main.saveFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetStatus(Lang.Get("main.statusSaved"));
            var answer = MessageBox.Show(this, Lang.Get("mail.saveSuccess"), Lang.Get("common.confirm"),
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (answer == DialogResult.Yes)
            {
                await RestartServiceAsync();
            }
        }

        private static void WriteAppSetting(XmlDocument doc, string key, string value)
        {
            XmlNode n = doc.SelectSingleNode("configuration/appSettings/add[@key='" + key + "']");
            if (n != null && n.Attributes["value"] != null)
            {
                n.Attributes["value"].Value = value;
            }
        }

        private async void BtnTestSend_Click(object sender, EventArgs e)
        {
            string recipient = InputBox.Show(this, Lang.Get("mail.testInputTitle"), Lang.Get("mail.testInputPrompt"), _txtSendTo.Text.Trim());
            if (recipient == null) return;

            string host = _txtMailHost.Text.Trim();
            string address = _txtMailAddress.Text.Trim();
            string password = _txtMailPassword.Text;
            string display = _txtMailDisplayName.Text.Trim();
            string programName = _txtProgramName.Text.Trim();
            bool ssl = _cboSsl.SelectedIndex == 0;

            var errors = new List<string>();
            if (host.Length == 0) errors.Add(Lang.Get("mail.missingHost"));
            if (address.Length == 0) errors.Add(Lang.Get("mail.missingAddress"));
            if (password.Length == 0) errors.Add(Lang.Get("mail.missingPassword"));
            int port;
            if (!int.TryParse(_txtPort.Text.Trim(), out port) || port <= 0) errors.Add(Lang.Get("mail.portInvalid"));
            if (!recipient.Contains("@")) errors.Add(Lang.Get("mail.invalidRecipient"));
            if (errors.Count > 0)
            {
                MessageBox.Show(this, string.Join(Environment.NewLine, errors), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string subject = Lang.Get("mail.testSubject", programName);
            string body = Lang.Get("mail.testBody", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            _btnTestSend.Enabled = false;
            SetStatus(Lang.Get("common.testing"));
            try
            {
                await Task.Run(() =>
                {
                    using (var client = new SmtpClient(host, port))
                    {
                        client.EnableSsl = ssl;
                        client.Credentials = new NetworkCredential(address, password);
                        client.Timeout = 30000;
                        using (var msg = new MailMessage())
                        {
                            msg.From = new MailAddress(address, display.Length > 0 ? display : address);
                            msg.To.Add(recipient);
                            msg.Subject = subject;
                            msg.Body = body;
                            client.Send(msg);
                        }
                    }
                });
                MessageBox.Show(this, Lang.Get("mail.testSuccess"), Lang.Get("common.success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetStatus(Lang.Get("mail.testSuccess"));
            }
            catch (Exception ex)
            {
                string detail = ex.GetBaseException().Message;
                MessageBox.Show(this, Lang.Get("mail.testFailed", detail), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus(Lang.Get("mail.testFailed", detail));
            }
            finally
            {
                _btnTestSend.Enabled = true;
            }
        }

        // ================= 服务状态/重启 =================
        private void RefreshServiceStatus()
        {
            var info = ServiceManager.GetStatus();
            switch (info.State)
            {
                case ServiceState.Running:
                    _lblServiceStatus.Text = Lang.Get("svc.statusRunning");
                    _lblServiceStatus.ForeColor = Color.ForestGreen;
                    break;
                case ServiceState.Stopped:
                    _lblServiceStatus.Text = Lang.Get("svc.statusStopped");
                    _lblServiceStatus.ForeColor = Color.Firebrick;
                    break;
                case ServiceState.NotInstalled:
                    _lblServiceStatus.Text = Lang.Get("common.notInstalled");
                    _lblServiceStatus.ForeColor = Color.Gray;
                    break;
                default:
                    _lblServiceStatus.Text = Lang.Get("svc.statusOther", info.RawStatus);
                    _lblServiceStatus.ForeColor = Color.DarkOrange;
                    break;
            }
            _btnRestart.Enabled = info.State != ServiceState.NotInstalled;
        }

        private async Task RestartServiceAsync()
        {
            var confirm = MessageBox.Show(this, Lang.Get("svc.restartConfirm"), Lang.Get("common.confirm"),
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (confirm != DialogResult.Yes) return;

            _btnRestart.Enabled = false;
            _btnRefreshStatus.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;
            SetStatus(Lang.Get("svc.restarting"));
            try
            {
                await Task.Run(() => ServiceManager.Restart());
                MessageBox.Show(this, Lang.Get("svc.restartSuccess"), Lang.Get("common.success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetStatus(Lang.Get("svc.restartSuccess"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus(Lang.Get("svc.restartFailed", ex.Message));
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                _btnRefreshStatus.Enabled = true;
                RefreshServiceStatus();
            }
        }

        // ================= 根目录/语言/刷新 =================
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.SelectedPath = _rootDir;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    ApplyRootDir(dlg.SelectedPath);
                }
            }
        }

        private void ApplyRootDir(string path)
        {
            path = (path ?? "").Trim();
            if (path.Length == 0) return;
            if (!Directory.Exists(path))
            {
                MessageBox.Show(this, Lang.Get("main.dirNotFound", path), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _rootDir = path;
            _txtRootDir.Text = path;
            SettingsStore.SaveRootDir(path);
            RefreshAll();
        }

        private void CboLang_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_loadingLangCombo) return;
            var info = _cboLang.SelectedItem as LangInfo;
            if (info == null || info.Code == Lang.CurrentCode) return;
            Lang.Load(info.Code);
            SettingsStore.SaveLanguage(info.Code);
            ApplyLanguageToAllForms();
        }

        private void ApplyLanguageToAllForms()
        {
            var forms = new List<Form>();
            foreach (Form f in Application.OpenForms) forms.Add(f);
            foreach (Form f in forms)
            {
                var loc = f as ILocalizableForm;
                if (loc != null) loc.ApplyLanguage();
            }
        }

        private void RefreshAll()
        {
            RefreshJobs();
            RefreshLogTree();
            LoadMailSettings();
            RefreshServiceStatus();

            bool hasConfig = File.Exists(ConfigModel.ServiceConfigPath(_rootDir));
            SetStatus(hasConfig
                ? Lang.Get("main.detectSummaryWithConfig", _grid.Rows.Count)
                : Lang.Get("main.detectSummaryNoConfig", _grid.Rows.Count));
        }

        private void SetStatus(string text)
        {
            _statusLabel.Text = text;
        }

        // ================= 多语言 =================
        public void ApplyLanguage()
        {
            Text = Lang.Get("main.title");

            _lblRootDir.Text = Lang.Get("main.rootDir");
            _btnBrowse.Text = Lang.Get("main.btnBrowse");
            _btnRefresh.Text = Lang.Get("common.refresh");
            _lblLang.Text = Lang.Get("main.langLabel");
            _lblServiceStatusTitle.Text = Lang.Get("main.serviceStatus");
            _btnRefreshStatus.Text = Lang.Get("main.btnRefreshStatus");
            _btnRestart.Text = Lang.Get("main.btnRestartService");

            _loadingLangCombo = true;
            string current = Lang.CurrentCode;
            _cboLang.Items.Clear();
            foreach (var info in Lang.GetAvailableLanguages())
            {
                _cboLang.Items.Add(info);
                if (info.Code == current) _cboLang.SelectedItem = info;
            }
            if (_cboLang.SelectedIndex < 0 && _cboLang.Items.Count > 0) _cboLang.SelectedIndex = 0;
            _loadingLangCombo = false;

            _tabJobs.Text = Lang.Get("main.tabJobs");
            _tabLogs.Text = Lang.Get("main.tabLogs");
            _tabMail.Text = Lang.Get("main.tabMail");

            _grid.Columns["NameCN"].HeaderText = Lang.Get("main.colNameCN");
            _grid.Columns["NameEN"].HeaderText = Lang.Get("main.colNameEN");
            _grid.Columns["Type"].HeaderText = Lang.Get("main.colType");
            _grid.Columns["Status"].HeaderText = Lang.Get("main.colStatus");
            _grid.Columns["Folder"].HeaderText = Lang.Get("main.colFolder");
            _grid.Columns["Setup"].HeaderText = Lang.Get("main.colSetup");
            _grid.Columns["Sequence"].HeaderText = Lang.Get("main.colSequence");
            _grid.Columns["Issue"].HeaderText = Lang.Get("main.colIssue");
            _btnNewJob.Text = Lang.Get("main.btnNewJob");
            _btnEdit.Text = Lang.Get("main.btnEdit");
            _btnEnable.Text = Lang.Get("main.btnEnable");
            _btnDisable.Text = Lang.Get("main.btnDisable");
            _btnOpenFolder.Text = Lang.Get("main.btnOpenFolder");

            _btnRefreshLogs.Text = Lang.Get("common.refresh");
            _txtPreview.Text = "";

            _lblMailNotice.Text = Lang.Get("mail.restartNotice");
            _lblMailHost.Text = Lang.Get("mail.host");
            _lblPort.Text = Lang.Get("mail.port");
            _lblMailAddress.Text = Lang.Get("mail.address");
            _lblMailDisplayName.Text = Lang.Get("mail.displayName");
            _lblMailPassword.Text = Lang.Get("mail.password");
            _chkShowPassword.Text = Lang.Get("mail.showPassword");
            _lblSsl.Text = Lang.Get("mail.ssl");
            _lblProgramName.Text = Lang.Get("mail.programName");
            _lblSendTo.Text = Lang.Get("mail.sendTo");
            _lblInfoMessage.Text = Lang.Get("mail.infoMessage");
            _btnSaveMail.Text = Lang.Get("common.save");
            _btnTestSend.Text = Lang.Get("mail.btnTestSend");

            RefreshAll();
        }
    }
}
