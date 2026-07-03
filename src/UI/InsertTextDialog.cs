using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class InsertTextDialog : Form
    {
        private readonly string _mode;
        private TextBox txtInput;
        private NumericUpDown nudPosition;
        private RadioButton rbFromStart;
        private RadioButton rbFromEnd;
        private Label lblHint;
        private Button btnOK;
        private Button btnCancel;

        public string InputText => txtInput.Text;
        public int? MidPosition => _mode == "mid" ? (int)nudPosition.Value : (int?)null;
        public bool MidFromEnd => _mode == "mid" && rbFromEnd.Checked;

        public InsertTextDialog(string mode)
        {
            _mode = mode ?? throw new ArgumentNullException(nameof(mode));

            string title;
            switch (_mode)
            {
                case "before": title = "前面添加文本"; break;
                case "after":  title = "末尾添加文本"; break;
                case "mid":    title = "中间插入文本"; break;
                default:       title = "插入文本"; break;
            }

            Text = title;
            Width = 420;
            int baseHeight = _mode == "mid" ? 340 : 230;
            Height = baseHeight;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            DialogStyleHelper.ApplyFormStyle(this);
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            int padding = 20;
            int y = 16;
            int fullWidth = Width - padding * 2 - 12;

            // ── 输入文本 ──────────────────────────────────────────────
            var lblInput = new Label
            {
                Text = "输入文本：",
                Location = new Point(padding, y),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblInput);
            Controls.Add(lblInput);
            y += 24;

            txtInput = new TextBox
            {
                Location = new Point(padding, y),
                Width = fullWidth,
                Height = 80,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            DialogStyleHelper.StyleTextBox(txtInput);
            Controls.Add(txtInput);
            y += 92;

            // ── 插入位置（仅 mid 模式）────────────────────────────────
            if (_mode == "mid")
            {
                // 插入方式
                var lblMode = new Label
                {
                    Text = "插入方式：",
                    Location = new Point(padding, y),
                    AutoSize = true
                };
                DialogStyleHelper.StyleLabel(lblMode);
                Controls.Add(lblMode);
                y += 24;

                rbFromStart = new RadioButton
                {
                    Text = "从开头",
                    Location = new Point(padding, y),
                    AutoSize = true,
                    Checked = true
                };
                DialogStyleHelper.StyleRadioButton(rbFromStart);
                Controls.Add(rbFromStart);

                rbFromEnd = new RadioButton
                {
                    Text = "从末尾",
                    Location = new Point(padding + 100, y),
                    AutoSize = true
                };
                DialogStyleHelper.StyleRadioButton(rbFromEnd);
                Controls.Add(rbFromEnd);

                y += 32;

                // 插入位置
                var lblPosition = new Label
                {
                    Text = "插入位置：",
                    Location = new Point(padding, y),
                    AutoSize = true
                };
                DialogStyleHelper.StyleLabel(lblPosition);
                Controls.Add(lblPosition);
                y += 24;

                nudPosition = new NumericUpDown
                {
                    Location = new Point(padding, y),
                    Width = 100,
                    Minimum = 1,
                    Maximum = 9999,
                    Value = 1
                };
                DialogStyleHelper.StyleNumericUpDown(nudPosition);
                Controls.Add(nudPosition);

                lblHint = new Label
                {
                    Text = "从开头第几个字符后插入",
                    Location = new Point(padding + 110, y + 4),
                    AutoSize = true,
                    ForeColor = DialogStyleHelper.TextSecondary,
                    Font = DialogStyleHelper.HeaderDescFont
                };
                Controls.Add(lblHint);

                // 切换方向时更新提示文字
                rbFromStart.CheckedChanged += (s, e) =>
                {
                    if (rbFromStart.Checked)
                        lblHint.Text = "从开头第几个字符后插入";
                };
                rbFromEnd.CheckedChanged += (s, e) =>
                {
                    if (rbFromEnd.Checked)
                        lblHint.Text = "从末尾第几个字符前插入";
                };

                y += 36;
            }

            // ── 底部按钮栏 ────────────────────────────────────────────
            var btnPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 56,
                BackColor = Color.FromArgb(248, 249, 250)
            };
            btnPanel.Paint += (s, e) =>
            {
                using (var pen = new Pen(DialogStyleHelper.HeaderBorder))
                    e.Graphics.DrawLine(pen, 0, 0, btnPanel.Width, 0);
            };
            Controls.Add(btnPanel);
            btnPanel.BringToFront();

            btnOK = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Width = 88,
                Height = 32,
                Location = new Point(btnPanel.Width - 196, 12)
            };
            DialogStyleHelper.StylePrimaryButton(btnOK);
            btnOK.Click += BtnOK_Click;
            btnPanel.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Width = 88,
                Height = 32,
                Location = new Point(btnPanel.Width - 98, 12)
            };
            DialogStyleHelper.StyleSecondaryButton(btnCancel);
            btnPanel.Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtInput.Text))
            {
                MessageBox.Show("请输入文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        }
    }
}
