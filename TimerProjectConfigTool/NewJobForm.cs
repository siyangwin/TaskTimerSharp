using System;
using System.Drawing;
using System.Windows.Forms;

namespace TimerProjectConfigTool
{
    public class NewJobForm : Form
    {
        private readonly string _rootDir;

        private Label _lblNameCN;
        private Label _lblNameEN;
        private Label _lblNameENHint;
        private Label _lblType;
        private TextBox _txtNameCN;
        private TextBox _txtNameEN;
        private RadioButton _rdoType0;
        private RadioButton _rdoType1;
        private RadioButton _rdoType2;
        private CheckBox _chkEnabled;
        private Button _btnOk;
        private Button _btnCancel;

        public string CreatedNameEN { get; private set; }

        public NewJobForm(string rootDir)
        {
            _rootDir = rootDir;
            BuildUi();
            ApplyLanguage();
        }

        private void BuildUi()
        {
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(460, 330);
            ShowInTaskbar = false;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                Padding = new Padding(12)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            _lblNameCN = FieldLabel();
            _txtNameCN = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) };
            _lblNameEN = FieldLabel();
            _txtNameEN = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) };
            _lblNameENHint = new Label { AutoSize = true, ForeColor = Color.Gray, Margin = new Padding(3, 0, 3, 6) };

            _lblType = FieldLabel();
            _rdoType0 = new RadioButton { AutoSize = true, Checked = true, Margin = new Padding(3, 8, 12, 3) };
            _rdoType1 = new RadioButton { AutoSize = true, Margin = new Padding(3, 8, 12, 3) };
            _rdoType2 = new RadioButton { AutoSize = true, Margin = new Padding(3, 8, 3, 3) };
            var typePanel = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0) };
            typePanel.Controls.Add(_rdoType0);
            typePanel.Controls.Add(_rdoType1);
            typePanel.Controls.Add(_rdoType2);

            _chkEnabled = new CheckBox { AutoSize = true, Checked = false, Margin = new Padding(3, 10, 3, 3) };

            _btnOk = new Button { AutoSize = true, Margin = new Padding(3, 14, 8, 3) };
            _btnOk.Click += BtnOk_Click;
            _btnCancel = new Button { AutoSize = true, Margin = new Padding(3, 14, 3, 3) };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            buttons.Controls.Add(_btnCancel);
            buttons.Controls.Add(_btnOk);

            int row = 0;
            AddRow(layout, row++, _lblNameCN, _txtNameCN);
            AddRow(layout, row++, _lblNameEN, _txtNameEN);
            layout.Controls.Add(_lblNameENHint, 1, row++);
            AddRow(layout, row++, _lblType, typePanel);
            layout.Controls.Add(_chkEnabled, 1, row++);
            layout.Controls.Add(buttons, 1, row);

            Controls.Add(layout);
            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        private static Label FieldLabel()
        {
            return new Label { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 9, 10, 3) };
        }

        private static void AddRow(TableLayoutPanel layout, int row, Label label, Control control)
        {
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(control, 1, row);
        }

        private void ApplyLanguage()
        {
            Text = Lang.Get("newjob.title");
            _lblNameCN.Text = Lang.Get("newjob.nameCN");
            _lblNameEN.Text = Lang.Get("newjob.nameEN");
            _lblNameENHint.Text = Lang.Get("newjob.nameENHint");
            _lblType.Text = Lang.Get("newjob.type");
            _rdoType0.Text = Lang.Get("newjob.typeApi");
            _rdoType1.Text = Lang.Get("newjob.typeSequence");
            _rdoType2.Text = Lang.Get("newjob.typeSp");
            _chkEnabled.Text = Lang.Get("newjob.status");
            _btnOk.Text = Lang.Get("newjob.btnCreate");
            _btnCancel.Text = Lang.Get("common.cancel");
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            string nameCN = _txtNameCN.Text.Trim();
            string nameEN = _txtNameEN.Text.Trim();
            string type = _rdoType1.Checked ? "1" : _rdoType2.Checked ? "2" : "0";

            if (nameCN.Length == 0)
            {
                Warn("newjob.nameCNRequired");
                return;
            }
            if (nameEN.Length == 0)
            {
                Warn("newjob.nameENRequired");
                return;
            }
            if (!ConfigModel.IsValidNameEN(nameEN))
            {
                Warn("newjob.nameENInvalid");
                return;
            }
            if (ConfigModel.NameENExists(_rootDir, nameEN))
            {
                Warn("newjob.nameENDuplicate");
                return;
            }

            try
            {
                ConfigModel.CreateJob(_rootDir, nameCN, nameEN, type, _chkEnabled.Checked);
                CreatedNameEN = nameEN;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Lang.Get("newjob.createFailed", ex.Message), Lang.Get("common.error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Warn(string key)
        {
            MessageBox.Show(this, Lang.Get(key), Lang.Get("common.warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
