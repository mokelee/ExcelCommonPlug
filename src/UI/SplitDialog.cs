using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class SplitDialog : Form
    {
        private NumericUpDown nudTitleRow;
        private TextBox txtInfoCol;
        private CheckBox chkSplitSave;
        private Button btnOK;
        private Button btnCancel;

        public int HeaderRows => (int)nudTitleRow.Value;
        public string SplitColumn => txtInfoCol.Text.Trim();
        public bool SaveAsFiles => chkSplitSave.Checked;

        public SplitDialog(int defaultTitleRow, string defaultInfoCol)
        {
            Text = "拆分工作表";
            Width = 360;
            Height = 220;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            DialogStyleHelper.ApplyFormStyle(this);
            InitializeComponent(defaultTitleRow, defaultInfoCol);
        }

        private void InitializeComponent(int defaultTitleRow, string defaultInfoCol)
        {
            int padding = 20;
            int labelWidth = 100;
            int y = 16;

            // 标题行数
            var lblTitleRow = new Label
            {
                Text = "标题行数：",
                Location = new Point(padding, y + 4),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblTitleRow);
            Controls.Add(lblTitleRow);

            nudTitleRow = new NumericUpDown
            {
                Location = new Point(padding + labelWidth, y),
                Width = 120,
                Minimum = 1,
                Maximum = 20,
                Value = Math.Max(1, defaultTitleRow)
            };
            DialogStyleHelper.StyleNumericUpDown(nudTitleRow);
            Controls.Add(nudTitleRow);

            y += 36;

            // 拆分列字母
            var lblInfoCol = new Label
            {
                Text = "拆分列字母：",
                Location = new Point(padding, y + 4),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblInfoCol);
            Controls.Add(lblInfoCol);

            txtInfoCol = new TextBox
            {
                Location = new Point(padding + labelWidth, y),
                Width = 120,
                Text = defaultInfoCol
            };
            DialogStyleHelper.StyleTextBox(txtInfoCol);
            Controls.Add(txtInfoCol);

            y += 36;

            // 另存为独立文件
            chkSplitSave = new CheckBox
            {
                Text = "同时另存为独立文件",
                Location = new Point(padding + labelWidth, y),
                AutoSize = true,
                Checked = false
            };
            DialogStyleHelper.StyleCheckBox(chkSplitSave);
            Controls.Add(chkSplitSave);

            // 底部按钮
            var btnPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
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
                Width = 80,
                Height = 30,
                Location = new Point(btnPanel.Width - 176, 10)
            };
            DialogStyleHelper.StylePrimaryButton(btnOK);
            btnOK.Click += BtnOK_Click;
            btnPanel.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Width = 80,
                Height = 30,
                Location = new Point(btnPanel.Width - 88, 10)
            };
            DialogStyleHelper.StyleSecondaryButton(btnCancel);
            btnPanel.Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtInfoCol.Text))
            {
                MessageBox.Show("请输入拆分列字母。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        }
    }
}
