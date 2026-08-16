using System.Drawing;
using System.Windows.Forms;

namespace TimerProjectConfigTool
{
    /// <summary>简单输入框对话框（用于测试发送填写收件人等场景），取消返回 null</summary>
    public static class InputBox
    {
        public static string Show(IWin32Window owner, string title, string prompt, string defaultValue)
        {
            using (Form form = new Form())
            {
                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    Padding = new Padding(12),
                    AutoSize = false
                };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var lbl = new Label
                {
                    Text = prompt,
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    Margin = new Padding(3, 3, 3, 8)
                };

                var txt = new TextBox
                {
                    Text = defaultValue ?? "",
                    Dock = DockStyle.Fill,
                    Margin = new Padding(3, 0, 3, 12)
                };

                var ok = new Button
                {
                    Text = Lang.Get("common.ok"),
                    DialogResult = DialogResult.OK,
                    Size = new Size(90, 28),
                    Margin = new Padding(3, 3, 8, 3)
                };
                var cancel = new Button
                {
                    Text = Lang.Get("common.cancel"),
                    DialogResult = DialogResult.Cancel,
                    Size = new Size(90, 28),
                    Margin = new Padding(3)
                };

                var buttons = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.RightToLeft,
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    Margin = new Padding(0)
                };
                buttons.Controls.Add(cancel);
                buttons.Controls.Add(ok);

                layout.Controls.Add(lbl, 0, 0);
                layout.Controls.Add(txt, 0, 1);
                layout.Controls.Add(buttons, 0, 2);

                form.Controls.Add(layout);
                form.Text = title;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.StartPosition = FormStartPosition.CenterParent;
                form.ClientSize = new Size(440, 130);
                form.AcceptButton = ok;
                form.CancelButton = cancel;
                form.ShowInTaskbar = false;

                return form.ShowDialog(owner) == DialogResult.OK ? txt.Text.Trim() : null;
            }
        }
    }
}
