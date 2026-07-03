using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class ExportImageDialog : Form
    {
        private NumericUpDown nudNameCol;
        private TextBox txtFolderPath;
        private Button btnBrowse;
        private RadioButton rbPng;
        private RadioButton rbJpg;
        private Button btnOK;
        private Button btnCancel;

        public int NameCol => (int)nudNameCol.Value;
        public string FolderPath => txtFolderPath.Text.Trim();
        public string PicExtension => rbJpg.Checked ? ".jpg" : ".png";

        public ExportImageDialog(int defaultNameCol)
        {
            Text = "另存图片";
            Width = 460;
            Height = 320;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            DialogStyleHelper.ApplyFormStyle(this);
            int startY = 16;
            InitializeComponent(startY, defaultNameCol);
        }

        private void InitializeComponent(int startY, int defaultNameCol)
        {
            int padding = 20;
            int labelWidth = 100;
            int controlLeft = padding + labelWidth;
            int controlWidth = Width - controlLeft - padding - 12;
            int y = startY;

            // Settings group
            var grpSettings = new GroupBox
            {
                Text = "导出设置",
                Location = new Point(padding, y),
                Width = Width - padding * 2 - 12,
                Height = 130
            };
            DialogStyleHelper.StyleGroupBox(grpSettings);
            Controls.Add(grpSettings);

            int gy = 24;

            // Name column
            var lblNameCol = new Label { Text = "名称列：", Location = new Point(14, gy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblNameCol);
            grpSettings.Controls.Add(lblNameCol);
            nudNameCol = new NumericUpDown
            {
                Location = new Point(14 + labelWidth, gy),
                Width = 100,
                Minimum = 1,
                Maximum = 9999,
                Value = Math.Max(1, defaultNameCol)
            };
            DialogStyleHelper.StyleNumericUpDown(nudNameCol);
            grpSettings.Controls.Add(nudNameCol);
            gy += 38;

            // Folder path
            var lblFolder = new Label { Text = "保存文件夹：", Location = new Point(14, gy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblFolder);
            grpSettings.Controls.Add(lblFolder);

            int txtWidth = grpSettings.Width - 14 - labelWidth - 75;
            txtFolderPath = new TextBox
            {
                Location = new Point(14 + labelWidth, gy),
                Width = txtWidth
            };
            DialogStyleHelper.StyleTextBox(txtFolderPath);
            grpSettings.Controls.Add(txtFolderPath);

            btnBrowse = new Button
            {
                Text = "浏览...",
                Width = 65,
                Height = 23,
                Location = new Point(14 + labelWidth + txtWidth + 6, gy)
            };
            DialogStyleHelper.StyleSmallButton(btnBrowse);
            btnBrowse.Click += BtnBrowse_Click;
            grpSettings.Controls.Add(btnBrowse);
            gy += 38;

            // Picture type
            var lblPicType = new Label { Text = "图片格式：", Location = new Point(14, gy + 4), AutoSize = true };
            DialogStyleHelper.StyleLabel(lblPicType);
            grpSettings.Controls.Add(lblPicType);

            rbPng = new RadioButton
            {
                Text = ".png",
                Location = new Point(14 + labelWidth, gy),
                AutoSize = true,
                Checked = true
            };
            DialogStyleHelper.StyleRadioButton(rbPng);
            grpSettings.Controls.Add(rbPng);

            rbJpg = new RadioButton
            {
                Text = ".jpg",
                Location = new Point(14 + labelWidth + 70, gy),
                AutoSize = true
            };
            DialogStyleHelper.StyleRadioButton(rbJpg);
            grpSettings.Controls.Add(rbJpg);

            y += 140;

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
                dlg.Description = "选择保存文件夹";
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
                MessageBox.Show("请选择保存文件夹。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        }
    }
}
