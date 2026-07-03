using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class CommentDialog : Form
    {
        private RadioButton rbRight;
        private RadioButton rbLeft;
        private NumericUpDown nudHGap;
        private RadioButton rbAbove;
        private RadioButton rbBelow;
        private NumericUpDown nudVGap;
        private NumericUpDown nudWidth;
        private NumericUpDown nudHeight;
        private RadioButton rbNoChange;
        private RadioButton rbDisplayAll;
        private RadioButton rbHideAll;
        private Button btnOK;
        private Button btnCancel;

        // 像素转磅: 1px = 0.75pt (96DPI)
        private const double PxToPt = 0.75;

        /// <summary>批注水平方向：true=右侧，false=左侧</summary>
        public bool IsRight => rbRight.Checked;

        /// <summary>批注垂直方向：true=上方，false=下方</summary>
        public bool IsAbove => rbAbove.Checked;

        /// <summary>水平间距（磅）</summary>
        public double HorizontalGap => (double)nudHGap.Value * PxToPt;

        /// <summary>垂直间距（磅）</summary>
        public double VerticalGap => (double)nudVGap.Value * PxToPt;

        /// <summary>批注宽度（磅）</summary>
        public double? CommentWidth => nudWidth.Enabled ? (double?)(nudWidth.Value * (decimal)PxToPt) : null;

        /// <summary>批注高度（磅）</summary>
        public double? CommentHeight => nudHeight.Enabled ? (double?)(nudHeight.Value * (decimal)PxToPt) : null;

        public string VisibilityMode
        {
            get
            {
                if (rbDisplayAll.Checked) return "show";
                if (rbHideAll.Checked) return "hide";
                return "none";
            }
        }

        public CommentDialog()
        {
            Text = "调整批注位置和大小";
            AutoScaleMode = AutoScaleMode.Dpi;
            Width = 580;
            Height = 460;
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
            int grpWidth = Width - padding * 2 - 16;

            // ========== 位置 GroupBox ==========
            var grpPos = new GroupBox
            {
                Text = "位置",
                Location = new Point(padding, y),
                Width = grpWidth,
                Height = 155
            };
            DialogStyleHelper.StyleGroupBox(grpPos);
            Controls.Add(grpPos);

            int gy = 28;
            int col1X = 20;
            int labelW = 85;
            int rbW = 88;
            int gapLabelX = col1X + labelW + rbW * 2 + 20;
            int gapLabelW = 52;
            int nudX = gapLabelX + gapLabelW;
            int nudW = 100;

            // --- Row 1: 水平方向 ---
            var lblH = new Label
            {
                Text = "水平方向：",
                Location = new Point(col1X, gy + 3),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblH);
            grpPos.Controls.Add(lblH);

            rbRight = new RadioButton
            {
                Text = "右侧",
                Location = new Point(col1X + labelW, gy),
                Width = rbW,
                Height = 24,
                Checked = true
            };
            DialogStyleHelper.StyleRadioButton(rbRight);
            grpPos.Controls.Add(rbRight);

            rbLeft = new RadioButton
            {
                Text = "左侧",
                Location = new Point(col1X + labelW + rbW, gy),
                Width = rbW,
                Height = 24
            };
            DialogStyleHelper.StyleRadioButton(rbLeft);
            grpPos.Controls.Add(rbLeft);

            var lblHGap = new Label
            {
                Text = "间距：",
                Location = new Point(gapLabelX, gy + 3),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblHGap);
            grpPos.Controls.Add(lblHGap);

            nudHGap = new NumericUpDown
            {
                Location = new Point(nudX, gy),
                Width = nudW,
                Height = 24,
                DecimalPlaces = 0,
                Minimum = 0,
                Maximum = 500,
                Value = 15,
                Increment = 5
            };
            DialogStyleHelper.StyleNumericUpDown(nudHGap);
            grpPos.Controls.Add(nudHGap);

            gy += 42;

            // --- Row 2: 垂直方向 ---
            var lblV = new Label
            {
                Text = "垂直方向：",
                Location = new Point(col1X, gy + 3),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblV);
            grpPos.Controls.Add(lblV);

            rbAbove = new RadioButton
            {
                Text = "上方",
                Location = new Point(col1X + labelW, gy),
                Width = rbW,
                Height = 24,
                Checked = true
            };
            DialogStyleHelper.StyleRadioButton(rbAbove);
            grpPos.Controls.Add(rbAbove);

            rbBelow = new RadioButton
            {
                Text = "下方",
                Location = new Point(col1X + labelW + rbW, gy),
                Width = rbW,
                Height = 24
            };
            DialogStyleHelper.StyleRadioButton(rbBelow);
            grpPos.Controls.Add(rbBelow);

            var lblVGap = new Label
            {
                Text = "间距：",
                Location = new Point(gapLabelX, gy + 3),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblVGap);
            grpPos.Controls.Add(lblVGap);

            nudVGap = new NumericUpDown
            {
                Location = new Point(nudX, gy),
                Width = nudW,
                Height = 24,
                DecimalPlaces = 0,
                Minimum = 0,
                Maximum = 500,
                Value = 10,
                Increment = 5
            };
            DialogStyleHelper.StyleNumericUpDown(nudVGap);
            grpPos.Controls.Add(nudVGap);

            gy += 42;

            // 备注
            var lblNote = new Label
            {
                Text = "间距：批注框边缘到单元格边缘的像素距离",
                Location = new Point(col1X, gy + 2),
                AutoSize = true,
                ForeColor = DialogStyleHelper.TextSecondary,
                Font = DialogStyleHelper.HeaderDescFont
            };
            grpPos.Controls.Add(lblNote);

            y += grpPos.Height + 12;

            // ========== 大小 GroupBox ==========
            var grpSize = new GroupBox
            {
                Text = "大小（像素）",
                Location = new Point(padding, y),
                Width = grpWidth,
                Height = 75
            };
            DialogStyleHelper.StyleGroupBox(grpSize);
            Controls.Add(grpSize);

            int sgy = 30;
            int sizeCol2X = grpSize.Width / 2;

            var lblWidth = new Label
            {
                Text = "宽度：",
                Location = new Point(col1X, sgy + 3),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblWidth);
            grpSize.Controls.Add(lblWidth);

            nudWidth = new NumericUpDown
            {
                Location = new Point(col1X + 56, sgy),
                Width = nudW,
                Height = 24,
                DecimalPlaces = 0,
                Minimum = 20,
                Maximum = 9999,
                Value = 150,
                Increment = 50
            };
            DialogStyleHelper.StyleNumericUpDown(nudWidth);
            grpSize.Controls.Add(nudWidth);

            var lblHeight = new Label
            {
                Text = "高度：",
                Location = new Point(sizeCol2X, sgy + 3),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblHeight);
            grpSize.Controls.Add(lblHeight);

            nudHeight = new NumericUpDown
            {
                Location = new Point(sizeCol2X + 56, sgy),
                Width = nudW,
                Height = 24,
                DecimalPlaces = 0,
                Minimum = 20,
                Maximum = 9999,
                Value = 100,
                Increment = 25
            };
            DialogStyleHelper.StyleNumericUpDown(nudHeight);
            grpSize.Controls.Add(nudHeight);

            y += grpSize.Height + 12;

            // ========== 可见性 GroupBox ==========
            var grpVis = new GroupBox
            {
                Text = "批注可见性",
                Location = new Point(padding, y),
                Width = grpWidth,
                Height = 64
            };
            DialogStyleHelper.StyleGroupBox(grpVis);
            Controls.Add(grpVis);

            rbNoChange = new RadioButton
            {
                Text = "不改变",
                Location = new Point(16, 28),
                Width = 80,
                Height = 24,
                Checked = true
            };
            DialogStyleHelper.StyleRadioButton(rbNoChange);
            grpVis.Controls.Add(rbNoChange);

            rbDisplayAll = new RadioButton
            {
                Text = "全部显示",
                Location = new Point(120, 28),
                Width = 90,
                Height = 24
            };
            DialogStyleHelper.StyleRadioButton(rbDisplayAll);
            grpVis.Controls.Add(rbDisplayAll);

            rbHideAll = new RadioButton
            {
                Text = "全部隐藏",
                Location = new Point(230, 28),
                Width = 90,
                Height = 24
            };
            DialogStyleHelper.StyleRadioButton(rbHideAll);
            grpVis.Controls.Add(rbHideAll);

            // ========== 底部按钮 ==========
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
    }
}
