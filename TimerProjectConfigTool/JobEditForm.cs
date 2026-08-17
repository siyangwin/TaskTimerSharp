using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TimerProjectConfigTool
{
    /// <summary>下拉项：Value 为写入 XML 的值，Display 为界面显示文字</summary>
    public class ComboItem
    {
        public string Value;
        public string Display;

        public ComboItem(string value, string display)
        {
            Value = value;
            Display = display;
        }

        public override string ToString()
        {
            return Display;
        }

        public static string SelectedValue(ComboBox cbo, string fallback)
        {
            var item = cbo.SelectedItem as ComboItem;
            return item != null ? item.Value : fallback;
        }

        public static void SelectByValue(ComboBox cbo, string value)
        {
            foreach (object o in cbo.Items)
            {
                var item = o as ComboItem;
                if (item != null && string.Equals(item.Value, value, StringComparison.OrdinalIgnoreCase))
                {
                    cbo.SelectedItem = item;
                    return;
                }
            }
            if (cbo.Items.Count > 0) cbo.SelectedIndex = 0;
        }
    }

    /// <summary>顺序步骤类型下拉项（Value 必须与服务端硬编码一致：API地址/存储过程）</summary>
    public class StepTypeItem
    {
        public string Value { get; set; }
        public string Display { get; set; }

        public override string ToString()
        {
            return Display;
        }
    }

    public class JobEditForm : Form, ILocalizableForm
    {
        private const string StepApi = "API地址";
        private const string StepSp = "存储过程";
        private static readonly Color HighlightColor = Color.FromArgb(255, 250, 220);

        private readonly string _rootDir;
        private readonly string _nameEN;

        private TabControl _tabs;
        private TabPage _tabBasic;
        private TabPage _tabSchedule;
        private TabPage _tabExecute;
        private TabPage _tabMail;
        private TabPage _tabSequence;

        // 基本信息
        private Label _lblNameCN;
        private TextBox _txtNameCN;
        private Label _lblType;
        private RadioButton _rdoType0;
        private RadioButton _rdoType1;
        private RadioButton _rdoType2;
        private CheckBox _chkEnabled;
        private Label _lblTypeHint;

        // 调度设置
        private Label _lblExecMode;
        private RadioButton _radioInterval;
        private RadioButton _radioCycle;
        private Label _lblIntervalsTime;
        private TextBox _txtIntervalsTime;
        private Label _lblCycleType;
        private ComboBox _cboCycleType;
        private Label _lblDayOfWeek;
        private ComboBox _cboDayOfWeek;
        private Label _lblSpecificTime;
        private TextBox _txtSpecificTime;
        private Label _lblStartTime;
        private TextBox _txtStartTime;
        private Label _lblEndTime;
        private TextBox _txtEndTime;
        private Label _lblTimeHint;

        // 执行设置
        private Label _lblType1Hint;
        private Label _lblApiUrl;
        private TextBox _txtApiUrl;
        private Label _lblConnStr;
        private TextBox _txtConnStr;
        private Button _btnTestConn;
        private Label _lblSp;
        private TextBox _txtStoredProcedure;
        private CheckBox _chkVerification;
        private Label _lblAuthKey;
        private TextBox _txtAuthKey;
        private Label _lblRetryCount;
        private TextBox _txtRetryCount;
        private Label _lblRetryInterval;
        private TextBox _txtRetryInterval;
        private Label _lblHttpTimeout;
        private TextBox _txtHttpTimeout;
        private Label _lblRemark;
        private TextBox _txtRemark;

        // 邮件通知
        private CheckBox _chkSendMail;
        private Label _lblMailTo;
        private TextBox _txtMailTo;

        // 顺序步骤
        private DataGridView _gridSteps;
        private Button _btnAddRow;
        private Button _btnDelRow;
        private Button _btnMoveUp;
        private Button _btnMoveDown;

        private Button _btnSave;
        private Button _btnCancel;

        private ToolTip _toolTip;

        public JobEditForm(string rootDir, string nameEN)
        {
            _rootDir = rootDir;
            _nameEN = nameEN;
            _toolTip = new ToolTip();
            BuildUi();
            LoadData();
            ApplyLanguage();
        }

        // ================= UI 构建 =================
        private void BuildUi()
        {
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(800, 680);
            ShowInTaskbar = false;

            _tabs = new TabControl { Dock = DockStyle.Fill };
            _tabBasic = new TabPage();
            _tabSchedule = new TabPage();
            _tabExecute = new TabPage();
            _tabMail = new TabPage();
            _tabSequence = new TabPage();
            _tabs.TabPages.Add(_tabBasic);
            _tabs.TabPages.Add(_tabSchedule);
            _tabs.TabPages.Add(_tabExecute);
            _tabs.TabPages.Add(_tabMail);
            _tabs.TabPages.Add(_tabSequence);

            BuildBasicTab();
            BuildScheduleTab();
            BuildExecuteTab();
            BuildMailTab();
            BuildSequenceTab();

            _btnSave = new Button { AutoSize = true, Margin = new Padding(3, 8, 8, 8) };
            _btnSave.Click += BtnSave_Click;
            _btnCancel = new Button { AutoSize = true, Margin = new Padding(3, 8, 3, 8) };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            var bottom = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                AutoSize = true
            };
            bottom.Controls.Add(_btnCancel);
            bottom.Controls.Add(_btnSave);

            Controls.Add(_tabs);
            Controls.Add(bottom);
            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private static TableLayoutPanel NewFieldTable()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoScroll = true,
                Padding = new Padding(10, 14, 10, 14)
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return t;
        }

        private static Label FieldLabel()
        {
            return new Label
            {
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(3, 9, 12, 3)
            };
        }

        private static TextBox FieldInput(int width)
        {
            return new TextBox { Width = width, Margin = new Padding(3, 6, 3, 8) };
        }

        private static void AddRow(TableLayoutPanel t, Label label, Control control)
        {
            int row = t.RowStyles.Count;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            if (label != null) t.Controls.Add(label, 0, row);
            t.Controls.Add(control, 1, row);
        }

        private static void AddFullRow(TableLayoutPanel t, Control control)
        {
            int row = t.RowStyles.Count;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(control, 0, row);
            t.SetColumnSpan(control, 2);
        }

        private void BuildBasicTab()
        {
            var t = NewFieldTable();

            _lblNameCN = FieldLabel();
            _txtNameCN = FieldInput(420);
            _lblType = FieldLabel();
            _rdoType0 = new RadioButton { AutoSize = true, Checked = true, Margin = new Padding(3, 6, 14, 3) };
            _rdoType1 = new RadioButton { AutoSize = true, Margin = new Padding(3, 6, 14, 3) };
            _rdoType2 = new RadioButton { AutoSize = true, Margin = new Padding(3, 6, 3, 3) };
            var typePanel = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0) };
            typePanel.Controls.Add(_rdoType0);
            typePanel.Controls.Add(_rdoType1);
            typePanel.Controls.Add(_rdoType2);
            _chkEnabled = new CheckBox { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            _lblTypeHint = new Label { AutoSize = false, Width = 560, ForeColor = Color.Gray, Margin = new Padding(3, 8, 3, 3) };
            _lblTypeHint.Height = 34;

            _rdoType0.CheckedChanged += (s, e) => UpdateTypeRelated();
            _rdoType1.CheckedChanged += (s, e) => UpdateTypeRelated();
            _rdoType2.CheckedChanged += (s, e) => UpdateTypeRelated();

            AddRow(t, _lblNameCN, _txtNameCN);
            AddRow(t, _lblType, typePanel);
            AddRow(t, null, _chkEnabled);
            AddFullRow(t, _lblTypeHint);
            _tabBasic.Controls.Add(t);
        }

        private void BuildScheduleTab()
        {
            var t = NewFieldTable();

            _lblExecMode = FieldLabel();
            _radioInterval = new RadioButton { AutoSize = true, Checked = true, Margin = new Padding(3, 6, 14, 3) };
            _radioCycle = new RadioButton { AutoSize = true, Margin = new Padding(3, 6, 3, 3) };
            var modePanel = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0) };
            modePanel.Controls.Add(_radioInterval);
            modePanel.Controls.Add(_radioCycle);

            _lblIntervalsTime = FieldLabel();
            _txtIntervalsTime = FieldInput(160);
            _lblCycleType = FieldLabel();
            _cboCycleType = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3, 6, 3, 3) };
            _lblDayOfWeek = FieldLabel();
            _cboDayOfWeek = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(3, 6, 3, 3) };
            _lblSpecificTime = FieldLabel();
            _txtSpecificTime = FieldInput(160);
            _lblStartTime = FieldLabel();
            _txtStartTime = FieldInput(220);
            _lblEndTime = FieldLabel();
            _txtEndTime = FieldInput(220);
            _lblTimeHint = new Label { AutoSize = true, ForeColor = Color.Gray, Margin = new Padding(3, 8, 3, 3) };

            _radioInterval.CheckedChanged += (s, e) => UpdateScheduleVisibility();
            _radioCycle.CheckedChanged += (s, e) => UpdateScheduleVisibility();
            _cboCycleType.SelectedIndexChanged += (s, e) => UpdateScheduleVisibility();

            AddRow(t, _lblExecMode, modePanel);
            AddRow(t, _lblIntervalsTime, _txtIntervalsTime);
            AddRow(t, _lblCycleType, _cboCycleType);
            AddRow(t, _lblDayOfWeek, _cboDayOfWeek);
            AddRow(t, _lblSpecificTime, _txtSpecificTime);
            AddRow(t, _lblStartTime, _txtStartTime);
            AddRow(t, _lblEndTime, _txtEndTime);
            AddFullRow(t, _lblTimeHint);
            _tabSchedule.Controls.Add(t);
        }

        private void BuildExecuteTab()
        {
            var t = NewFieldTable();

            _lblType1Hint = new Label { AutoSize = false, Width = 560, ForeColor = Color.Gray, Margin = new Padding(3, 3, 3, 3) };
            _lblType1Hint.Height = 20;
            _lblApiUrl = FieldLabel();
            _txtApiUrl = FieldInput(560);
            _lblConnStr = FieldLabel();
            _txtConnStr = FieldInput(430);
            _btnTestConn = new Button { AutoSize = true, Margin = new Padding(8, 3, 3, 3) };
            _btnTestConn.Click += BtnTestConn_Click;
            var connPanel = new FlowLayoutPanel { AutoSize = true, Margin = new Padding(0) };
            connPanel.Controls.Add(_txtConnStr);
            connPanel.Controls.Add(_btnTestConn);
            _lblSp = FieldLabel();
            _txtStoredProcedure = FieldInput(430);
            _chkVerification = new CheckBox { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            _chkVerification.CheckedChanged += (s, e) => UpdateExecuteVisibility();
            _lblAuthKey = FieldLabel();
            _txtAuthKey = FieldInput(430);
            _lblRetryCount = FieldLabel();
            _txtRetryCount = FieldInput(120);
            _lblRetryInterval = FieldLabel();
            _txtRetryInterval = FieldInput(120);
            _lblHttpTimeout = FieldLabel();
            _txtHttpTimeout = FieldInput(120);
            _lblRemark = FieldLabel();
            _txtRemark = FieldInput(560);
            _txtRemark.Multiline = true;
            _txtRemark.Height = 58;

            AddFullRow(t, _lblType1Hint);
            AddRow(t, _lblApiUrl, _txtApiUrl);
            AddRow(t, _lblConnStr, connPanel);
            AddRow(t, _lblSp, _txtStoredProcedure);
            AddRow(t, null, _chkVerification);
            AddRow(t, _lblAuthKey, _txtAuthKey);
            AddRow(t, _lblRetryCount, _txtRetryCount);
            AddRow(t, _lblRetryInterval, _txtRetryInterval);
            AddRow(t, _lblHttpTimeout, _txtHttpTimeout);
            AddRow(t, _lblRemark, _txtRemark);
            _tabExecute.Controls.Add(t);
        }

        private void BuildMailTab()
        {
            var t = NewFieldTable();

            _chkSendMail = new CheckBox { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            _lblMailTo = FieldLabel();
            _txtMailTo = FieldInput(560);
            _txtMailTo.Multiline = true;
            _txtMailTo.Height = 50;

            AddRow(t, null, _chkSendMail);
            AddRow(t, _lblMailTo, _txtMailTo);
            _tabMail.Controls.Add(t);
        }

        private void BuildSequenceTab()
        {
            _gridSteps = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.Fixed3D,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            var colType = new DataGridViewComboBoxColumn
            {
                Name = "StepType",
                Width = 150,
                FlatStyle = FlatStyle.Flat,
                DataSource = BuildStepTypes(),
                ValueMember = "Value",
                DisplayMember = "Display"
            };
            var colInfo = new DataGridViewTextBoxColumn { Name = "Info", Width = 560 };
            _gridSteps.Columns.Add(colType);
            _gridSteps.Columns.Add(colInfo);

            _btnAddRow = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnAddRow.Click += (s, e) => { _gridSteps.Rows.Add(StepApi, ""); _gridSteps.ClearSelection(); _gridSteps.Rows[_gridSteps.Rows.Count - 1].Selected = true; };
            _btnDelRow = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnDelRow.Click += BtnDelRow_Click;
            _btnMoveUp = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnMoveUp.Click += (s, e) => MoveStep(-1);
            _btnMoveDown = new Button { AutoSize = true, Margin = new Padding(3) };
            _btnMoveDown.Click += (s, e) => MoveStep(1);

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(4) };
            bar.Controls.Add(_btnAddRow);
            bar.Controls.Add(_btnDelRow);
            bar.Controls.Add(_btnMoveUp);
            bar.Controls.Add(_btnMoveDown);

            _tabSequence.Controls.Add(_gridSteps);
            _tabSequence.Controls.Add(bar);
        }

        private StepTypeItem[] BuildStepTypes()
        {
            return new[]
            {
                new StepTypeItem { Value = StepApi, Display = Lang.Get("jobedit.stepApi") },
                new StepTypeItem { Value = StepSp, Display = Lang.Get("jobedit.stepSp") }
            };
        }

        // ================= 数据加载 =================
        private void LoadData()
        {
            JobRegistryEntry entry = null;
            try
            {
                foreach (var e in ConfigModel.LoadRegistry(_rootDir))
                {
                    if (string.Equals(e.NameEN, _nameEN, StringComparison.OrdinalIgnoreCase))
                    {
                        entry = e;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("jobedit.loadFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            JobSetup setup;
            List<SequenceStep> steps;
            try
            {
                setup = ConfigModel.LoadSetup(_rootDir, _nameEN);
                steps = ConfigModel.LoadSequence(_rootDir, _nameEN);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("jobedit.loadFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                setup = new JobSetup();
                steps = new List<SequenceStep>();
            }

            _txtNameCN.Text = entry != null ? entry.NameCN : "";
            string type = entry != null ? entry.Type : "0";
            _rdoType0.Checked = type != "1" && type != "2";
            _rdoType1.Checked = type == "1";
            _rdoType2.Checked = type == "2";
            _chkEnabled.Checked = entry == null || entry.Enabled;

            _radioInterval.Checked = setup.ExecutionStatus != "1";
            _radioCycle.Checked = setup.ExecutionStatus == "1";
            _txtIntervalsTime.Text = setup.IntervalsTime;
            _txtSpecificTime.Text = setup.SpecificTime;
            _txtStartTime.Text = setup.StartTime;
            _txtEndTime.Text = setup.EndTime;

            _txtApiUrl.Text = setup.ApiUrl;
            _txtConnStr.Text = setup.ConnStr;
            _txtStoredProcedure.Text = setup.StoredProcedure;
            _chkVerification.Checked = string.Equals(setup.Verification, "true", StringComparison.OrdinalIgnoreCase);
            _txtAuthKey.Text = setup.AuthenticationKey;
            _txtRetryCount.Text = setup.RetryCount;
            _txtRetryInterval.Text = setup.RetryInterval;
            _txtHttpTimeout.Text = setup.HttpTimeout;
            _txtRemark.Text = setup.Remark;

            _chkSendMail.Checked = string.Equals(setup.SendMail, "true", StringComparison.OrdinalIgnoreCase);
            _txtMailTo.Text = setup.MailTo;

            _setupCyclePending = setup.CycltType;
            _setupDayPending = setup.DayOfWeek;

            foreach (var step in steps)
            {
                string project = step.Project == StepSp ? StepSp : StepApi;
                _gridSteps.Rows.Add(project, step.Info);
            }
        }

        private string _setupCyclePending;
        private string _setupDayPending;

        // ================= 联动显隐 =================
        private string SelectedType()
        {
            return _rdoType1.Checked ? "1" : _rdoType2.Checked ? "2" : "0";
        }

        private void UpdateTypeRelated()
        {
            if (_lblTypeHint == null) return;
            if (_rdoType0.Checked) _lblTypeHint.Text = Lang.Get("jobedit.typeHintApi");
            else if (_rdoType1.Checked) _lblTypeHint.Text = Lang.Get("jobedit.typeHintSequence");
            else _lblTypeHint.Text = Lang.Get("jobedit.typeHintSp");
            UpdateExecuteVisibility();
        }

        private void UpdateScheduleVisibility()
        {
            bool cycle = _radioCycle.Checked;
            _lblIntervalsTime.Visible = _txtIntervalsTime.Visible = !cycle;
            _lblCycleType.Visible = _cboCycleType.Visible = cycle;
            _lblSpecificTime.Visible = _txtSpecificTime.Visible = cycle;
            bool weekly = cycle && ComboItem.SelectedValue(_cboCycleType, "EveryDay") == "EveryWeek";
            _lblDayOfWeek.Visible = _cboDayOfWeek.Visible = weekly;
        }

        private void UpdateExecuteVisibility()
        {
            string type = SelectedType();
            _lblType1Hint.Visible = type == "1";
            _lblApiUrl.Visible = _txtApiUrl.Visible = type == "0";
            _lblSp.Visible = _txtStoredProcedure.Visible = type == "2";
            bool showKey = _chkVerification.Checked;
            _lblAuthKey.Visible = _txtAuthKey.Visible = showKey;

            _txtApiUrl.BackColor = type == "0" ? HighlightColor : SystemColors.Window;
            _txtConnStr.BackColor = type == "2" ? HighlightColor : SystemColors.Window;
            _txtStoredProcedure.BackColor = type == "2" ? HighlightColor : SystemColors.Window;
        }

        // ================= 顺序步骤操作 =================
        private void BtnDelRow_Click(object sender, EventArgs e)
        {
            if (_gridSteps.CurrentRow == null) return;
            _gridSteps.Rows.Remove(_gridSteps.CurrentRow);
        }

        private void MoveStep(int delta)
        {
            var row = _gridSteps.CurrentRow;
            if (row == null) return;
            int from = row.Index;
            int to = from + delta;
            if (to < 0 || to >= _gridSteps.Rows.Count) return;

            object t1 = _gridSteps.Rows[from].Cells["StepType"].Value;
            object i1 = _gridSteps.Rows[from].Cells["Info"].Value;
            _gridSteps.Rows[from].Cells["StepType"].Value = _gridSteps.Rows[to].Cells["StepType"].Value;
            _gridSteps.Rows[from].Cells["Info"].Value = _gridSteps.Rows[to].Cells["Info"].Value;
            _gridSteps.Rows[to].Cells["StepType"].Value = t1;
            _gridSteps.Rows[to].Cells["Info"].Value = i1;

            _gridSteps.ClearSelection();
            _gridSteps.Rows[to].Selected = true;
            _gridSteps.CurrentCell = _gridSteps.Rows[to].Cells["Info"];
        }

        private List<SequenceStep> CollectSteps()
        {
            var steps = new List<SequenceStep>();
            foreach (DataGridViewRow row in _gridSteps.Rows)
            {
                if (row.IsNewRow) continue;
                string project = Convert.ToString(row.Cells["StepType"].Value);
                if (project != StepSp) project = StepApi;
                steps.Add(new SequenceStep { Project = project, Info = Convert.ToString(row.Cells["Info"].Value).Trim() });
            }
            return steps;
        }

        // ================= 测试连接 =================
        private async void BtnTestConn_Click(object sender, EventArgs e)
        {
            string connStr = _txtConnStr.Text.Trim();
            if (connStr.Length == 0)
            {
                MessageBox.Show(this, Lang.Get("jobedit.testConnEmpty"), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _btnTestConn.Enabled = false;
            try
            {
                await Task.Run(() =>
                {
                    using (var conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                    }
                });
                MessageBox.Show(this, Lang.Get("jobedit.testConnSuccess"), Lang.Get("jobedit.testConnTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("jobedit.testConnFailed", ex.GetBaseException().Message), Lang.Get("jobedit.testConnTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnTestConn.Enabled = true;
            }
        }

        // ================= 校验与保存 =================
        private List<string> ValidateInput()
        {
            var errors = new List<string>();
            string type = SelectedType();

            if (type == "0" && _txtApiUrl.Text.Trim().Length == 0)
                errors.Add(Lang.Get("val.apiUrlRequired"));
            if (type == "2")
            {
                if (_txtConnStr.Text.Trim().Length == 0) errors.Add(Lang.Get("val.connStrRequired"));
                if (_txtStoredProcedure.Text.Trim().Length == 0) errors.Add(Lang.Get("val.spRequired"));
            }
            if (type == "1")
            {
                bool hasStep = false;
                foreach (var step in CollectSteps())
                {
                    if (step.Info.Length > 0) { hasStep = true; break; }
                }
                if (!hasStep) errors.Add(Lang.Get("val.stepsRequired"));
            }

            if (_radioCycle.Checked)
            {
                string cycle = ComboItem.SelectedValue(_cboCycleType, "");
                if (cycle != "EveryDay" && cycle != "EveryWeek") errors.Add(Lang.Get("val.cycleTypeInvalid"));
                DateTime t;
                if (!DateTime.TryParseExact(_txtSpecificTime.Text.Trim(), "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out t))
                    errors.Add(Lang.Get("val.specificTimeInvalid"));
                if (cycle == "EveryWeek")
                {
                    string dow = ComboItem.SelectedValue(_cboDayOfWeek, "");
                    if (!IsValidDayOfWeek(dow)) errors.Add(Lang.Get("val.dayOfWeekInvalid"));
                }
            }
            else
            {
                double interval;
                if (!double.TryParse(_txtIntervalsTime.Text.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out interval) || interval <= 0)
                    errors.Add(Lang.Get("val.intervalsInvalid"));
            }

            DateTime parsed;
            if (_txtStartTime.Text.Trim().Length > 0 &&
                !DateTime.TryParseExact(_txtStartTime.Text.Trim(), "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                errors.Add(Lang.Get("val.startTimeInvalid"));
            if (_txtEndTime.Text.Trim().Length > 0 &&
                !DateTime.TryParseExact(_txtEndTime.Text.Trim(), "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                errors.Add(Lang.Get("val.endTimeInvalid"));

            int retryCount;
            if (_txtRetryCount.Text.Trim().Length > 0 &&
                (!int.TryParse(_txtRetryCount.Text.Trim(), out retryCount) || retryCount < 0 || retryCount > 10))
                errors.Add(Lang.Get("val.retryCountInvalid"));
            int retryInterval;
            if (_txtRetryInterval.Text.Trim().Length > 0 &&
                (!int.TryParse(_txtRetryInterval.Text.Trim(), out retryInterval) || retryInterval < 0))
                errors.Add(Lang.Get("val.retryIntervalInvalid"));
            int httpTimeout;
            if (_txtHttpTimeout.Text.Trim().Length > 0 &&
                (!int.TryParse(_txtHttpTimeout.Text.Trim(), out httpTimeout) || httpTimeout <= 0))
                errors.Add(Lang.Get("val.httpTimeoutInvalid"));

            if (_chkSendMail.Checked && _txtMailTo.Text.Trim().Length > 0)
            {
                foreach (string part in _txtMailTo.Text.Split(';'))
                {
                    string addr = part.Trim();
                    if (addr.Length > 0 && !addr.Contains("@"))
                    {
                        errors.Add(Lang.Get("val.mailToInvalid"));
                        break;
                    }
                }
            }

            return errors;
        }

        private static bool IsValidDayOfWeek(string value)
        {
            switch (value)
            {
                case "Monday":
                case "Tuesday":
                case "Wednesday":
                case "Thursday":
                case "Friday":
                case "Saturday":
                case "Sunday":
                    return true;
                default:
                    return false;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var errors = ValidateInput();
            if (errors.Count > 0)
            {
                MessageBox.Show(this, string.Join(Environment.NewLine, errors), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string type = SelectedType();
            try
            {
                var setup = new JobSetup
                {
                    StartTime = _txtStartTime.Text.Trim(),
                    EndTime = _txtEndTime.Text.Trim(),
                    ExecutionStatus = _radioCycle.Checked ? "1" : "0",
                    IntervalsTime = _txtIntervalsTime.Text.Trim(),
                    CycltType = ComboItem.SelectedValue(_cboCycleType, "EveryDay"),
                    DayOfWeek = ComboItem.SelectedValue(_cboDayOfWeek, "Monday"),
                    SpecificTime = _txtSpecificTime.Text.Trim(),
                    ApiUrl = _txtApiUrl.Text.Trim(),
                    StoredProcedure = _txtStoredProcedure.Text.Trim(),
                    Remark = _txtRemark.Text.Trim(),
                    MailTo = _txtMailTo.Text.Trim(),
                    SendMail = _chkSendMail.Checked ? "True" : "False",
                    ConnStr = _txtConnStr.Text.Trim(),
                    Verification = _chkVerification.Checked ? "True" : "False",
                    AuthenticationKey = _txtAuthKey.Text.Trim(),
                    RetryCount = _txtRetryCount.Text.Trim().Length > 0 ? _txtRetryCount.Text.Trim() : "0",
                    RetryInterval = _txtRetryInterval.Text.Trim().Length > 0 ? _txtRetryInterval.Text.Trim() : "60",
                    HttpTimeout = _txtHttpTimeout.Text.Trim().Length > 0 ? _txtHttpTimeout.Text.Trim() : "600"
                };

                ConfigModel.SaveSetup(_rootDir, _nameEN, setup);
                if (type == "1")
                {
                    ConfigModel.SaveSequence(_rootDir, _nameEN, CollectSteps());
                }

                var registry = ConfigModel.LoadRegistry(_rootDir);
                JobRegistryEntry entry = null;
                foreach (var item in registry)
                {
                    if (string.Equals(item.NameEN, _nameEN, StringComparison.OrdinalIgnoreCase))
                    {
                        entry = item;
                        break;
                    }
                }
                if (entry == null)
                {
                    entry = new JobRegistryEntry { NameEN = _nameEN };
                    registry.Add(entry);
                }
                entry.NameCN = _txtNameCN.Text.Trim();
                entry.Type = type;
                entry.Enabled = _chkEnabled.Checked;
                ConfigModel.SaveRegistry(_rootDir, registry);

                MessageBox.Show(this, Lang.Get("jobedit.saveSuccess"), Lang.Get("common.success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("jobedit.saveFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================= 多语言 =================
        public void ApplyLanguage()
        {
            Text = Lang.Get("jobedit.title", _nameEN);

            _tabBasic.Text = Lang.Get("jobedit.tabBasic");
            _tabSchedule.Text = Lang.Get("jobedit.tabSchedule");
            _tabExecute.Text = Lang.Get("jobedit.tabExecute");
            _tabMail.Text = Lang.Get("jobedit.tabMail");
            _tabSequence.Text = Lang.Get("jobedit.tabSequence");

            _lblNameCN.Text = Lang.Get("jobedit.nameCN");
            _lblType.Text = Lang.Get("jobedit.type");
            _rdoType0.Text = Lang.Get("newjob.typeApi");
            _rdoType1.Text = Lang.Get("newjob.typeSequence");
            _rdoType2.Text = Lang.Get("newjob.typeSp");
            _chkEnabled.Text = Lang.Get("jobedit.enabled");

            _lblExecMode.Text = Lang.Get("jobedit.execMode");
            _radioInterval.Text = Lang.Get("jobedit.modeInterval");
            _radioCycle.Text = Lang.Get("jobedit.modeCycle");
            _lblIntervalsTime.Text = Lang.Get("jobedit.intervalsTime");
            _lblCycleType.Text = Lang.Get("jobedit.cycleType");
            _lblDayOfWeek.Text = Lang.Get("jobedit.dayOfWeek");
            _lblSpecificTime.Text = Lang.Get("jobedit.specificTime");
            _lblStartTime.Text = Lang.Get("jobedit.startTime");
            _lblEndTime.Text = Lang.Get("jobedit.endTime");
            _lblTimeHint.Text = Lang.Get("jobedit.timeFormatHint");

            string cycleSelected = ComboItem.SelectedValue(_cboCycleType, _setupCyclePending);
            string daySelected = ComboItem.SelectedValue(_cboDayOfWeek, _setupDayPending);
            _cboCycleType.Items.Clear();
            _cboCycleType.Items.Add(new ComboItem("EveryDay", Lang.Get("jobedit.everyDay")));
            _cboCycleType.Items.Add(new ComboItem("EveryWeek", Lang.Get("jobedit.everyWeek")));
            ComboItem.SelectByValue(_cboCycleType, cycleSelected.Length > 0 ? cycleSelected : "EveryDay");

            _cboDayOfWeek.Items.Clear();
            string[] days = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
            foreach (string d in days)
            {
                _cboDayOfWeek.Items.Add(new ComboItem(d, Lang.Get("dow." + d)));
            }
            ComboItem.SelectByValue(_cboDayOfWeek, daySelected.Length > 0 ? daySelected : "Monday");
            _setupCyclePending = null;
            _setupDayPending = null;

            _lblType1Hint.Text = Lang.Get("jobedit.typeHintSequence");
            _lblApiUrl.Text = Lang.Get("jobedit.apiUrl");
            _lblConnStr.Text = Lang.Get("jobedit.connStr");
            _btnTestConn.Text = Lang.Get("jobedit.btnTestConn");
            _lblSp.Text = Lang.Get("jobedit.storedProcedure");
            _chkVerification.Text = Lang.Get("jobedit.verification");
            _lblAuthKey.Text = Lang.Get("jobedit.authKey");
            _lblRetryCount.Text = Lang.Get("jobedit.retryCount");
            _toolTip.SetToolTip(_lblRetryCount, Lang.Get("jobedit.retrySpWarning"));
            _lblRetryInterval.Text = Lang.Get("jobedit.retryInterval");
            _lblHttpTimeout.Text = Lang.Get("jobedit.httpTimeout");
            _lblRemark.Text = Lang.Get("jobedit.remark");

            _chkSendMail.Text = Lang.Get("jobedit.sendMail");
            _lblMailTo.Text = Lang.Get("jobedit.mailTo");

            var colType = _gridSteps.Columns["StepType"] as DataGridViewComboBoxColumn;
            if (colType != null)
            {
                colType.DataSource = BuildStepTypes();
                colType.HeaderText = Lang.Get("jobedit.stepType");
            }
            _gridSteps.Columns["Info"].HeaderText = Lang.Get("jobedit.stepContent");
            _btnAddRow.Text = Lang.Get("jobedit.btnAddRow");
            _btnDelRow.Text = Lang.Get("jobedit.btnDelRow");
            _btnMoveUp.Text = Lang.Get("jobedit.btnMoveUp");
            _btnMoveDown.Text = Lang.Get("jobedit.btnMoveDown");

            _btnSave.Text = Lang.Get("common.save");
            _btnCancel.Text = Lang.Get("common.cancel");

            UpdateTypeRelated();
            UpdateScheduleVisibility();
            UpdateExecuteVisibility();
        }
    }
}
