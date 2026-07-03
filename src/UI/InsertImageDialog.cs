using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class InsertImageDialog : Form
    {
        private NumericUpDown nudLeftCol;
        private NumericUpDown nudRightCol;
        private NumericUpDown nudStartRow;
        private TextBox txtFolderPath;
        private Button btnBrowse;
        private RadioButton rbPng;
        private RadioButton rbJpg;
        private Button btnOK;
        private Button btnCancel;

        public int LeftCol => (int)nudLeftCol.Value;
        public int RightCol => (int)nudRightCol.Value;
        public int StartRow => (int)nudStartRow.Value;
        public string FolderPath => txtFolderPath.Text.Trim();
        public string PicExtension => rbJpg.Checked ? ".jpg" : ".png";

        public InsertImageDialog(int defaultLeftCol, int defaultRightCol)
        {
            Text = "插入图片";
            Width = 480;
            Height = 380;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            DialogStyleHelper.ApplyFormStyle(this);
            int startY = 16;
            InitializeComponent(startY, defaultLeftCol, defaultRightCol);
        }

        private void InitializeComponent(int startY, int defaultLeftCol, int defaultRightCol)
        {
            int padding = 20;
            int labelWidth = 105;
            int controlLeft = padding + labelWidth;
            int y = startY;

            // Position settings group
            var grpPos = new GroupBox
            {
                Text = "位置设置",
                Location = new Point(padding, y),
                Width = Width - padding * 2 - 12,
                Height = 120
            };
            DialogStyleHelper.StyleGroupBox(grpPos);
            Controls.Add(grpPos);

            int gy = 24;

            // Left column
            var lblLeft = new Label { Text = "名称列（左）：", Location = new Point(14, gy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblLeft);
            grpPos.Controls.Add(lblLeft);
            nudLeftCol = new NumericUpDown
            {
                Location = new Point(14 + labelWidth, gy),
                Width = 100,
                Minimum = 1,
                Maximum = 9999,
                Value = Math.Max(1, defaultLeftCol)
            };
            DialogStyleHelper.StyleNumericUpDown(nudLeftCol);
            grpPos.Controls.Add(nudLeftCol);
            gy += 34;

            // Right column
            var lblRight = new Label { Text = "图片列（右）：", Location = new Point(14, gy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblRight);
            grpPos.Controls.Add(lblRight);
            nudRightCol = new NumericUpDown
            {
                Location = new Point(14 + labelWidth, gy),
                Width = 100,
                Minimum = 1,
                Maximum = 9999,
                Value = Math.Max(1, defaultRightCol)
            };
            DialogStyleHelper.StyleNumericUpDown(nudRightCol);
            grpPos.Controls.Add(nudRightCol);
            gy += 34;

            // Start row
            var lblStartRow = new Label { Text = "起始行：", Location = new Point(14, gy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblStartRow);
            grpPos.Controls.Add(lblStartRow);
            nudStartRow = new NumericUpDown
            {
                Location = new Point(14 + labelWidth, gy),
                Width = 100,
                Minimum = 1,
                Maximum = 99999,
                Value = 2
            };
            DialogStyleHelper.StyleNumericUpDown(nudStartRow);
            grpPos.Controls.Add(nudStartRow);

            y += 130;

            // File settings group
            var grpFile = new GroupBox
            {
                Text = "文件设置",
                Location = new Point(padding, y),
                Width = Width - padding * 2 - 12,
                Height = 96
            };
            DialogStyleHelper.StyleGroupBox(grpFile);
            Controls.Add(grpFile);

            // Folder path
            int fy = 24;
            var lblFolder = new Label { Text = "图片文件夹：", Location = new Point(14, fy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblFolder);
            grpFile.Controls.Add(lblFolder);

            int txtWidth = grpFile.Width - 14 - labelWidth - 75;
            txtFolderPath = new TextBox
            {
                Location = new Point(14 + labelWidth, fy),
                Width = txtWidth
            };
            DialogStyleHelper.StyleTextBox(txtFolderPath);
            grpFile.Controls.Add(txtFolderPath);

            btnBrowse = new Button
            {
                Text = "浏览...",
                Width = 65,
                Height = 23,
                Location = new Point(14 + labelWidth + txtWidth + 6, fy)
            };
            DialogStyleHelper.StyleSmallButton(btnBrowse);
            btnBrowse.Click += BtnBrowse_Click;
            grpFile.Controls.Add(btnBrowse);

            fy += 36;

            // Picture type
            var lblPicType = new Label { Text = "图片格式：", Location = new Point(14, fy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblPicType);
            grpFile.Controls.Add(lblPicType);

            rbPng = new RadioButton
            {
                Text = ".png",
                Location = new Point(14 + labelWidth, fy),
                AutoSize = true,
                Checked = true
            };
            DialogStyleHelper.StyleRadioButton(rbPng);
            grpFile.Controls.Add(rbPng);

            rbJpg = new RadioButton
            {
                Text = ".jpg",
                Location = new Point(14 + labelWidth + 70, fy),
                AutoSize = true
            };
            DialogStyleHelper.StyleRadioButton(rbJpg);
            grpFile.Controls.Add(rbJpg);

            y += 106;

            // Bottom button panel
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

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "选择图片文件夹";
                if (!string.IsNullOrEmpty(txtFolderPath.Text))
                    dlg.SelectedPath = txtFolderPath.Text;
                if (dlg.ShowDialog() == DialogResult.OK)
                    txtFolderPath.Text = dlg.SelectedPath;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("请选择图片文件夹。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        }
    }
}
