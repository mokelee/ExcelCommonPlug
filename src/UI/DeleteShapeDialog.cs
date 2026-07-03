using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class DeleteShapeDialog : Form
    {
        private static readonly Dictionary<string, int> ShapeTypeMap = new Dictionary<string, int>
        {
            { "批注",     4 },   // msoComment
            { "图片",    13 },   // msoPicture
            { "文本框",  17 },   // msoTextBox
            { "SmartArt", 25 },  // msoSmartArt (Office 2010+)
            { "图表",     3 },   // msoChart
            { "形状",     1 }    // msoAutoShape
        };

        private CheckedListBox clbShapeTypes;
        private Button btnSelectAll;
        private Button btnUnselectAll;
        private Button btnOK;
        private Button btnCancel;

        public List<int> SelectedShapeTypes
        {
            get
            {
                var result = new List<int>();
                foreach (var item in clbShapeTypes.CheckedItems)
                {
                    string name = item.ToString();
                    if (ShapeTypeMap.TryGetValue(name, out int type))
                        result.Add(type);
                }
                return result;
            }
        }

        public DeleteShapeDialog()
        {
            Text = "删除图形";
            Width = 300;
            Height = 320;
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
            int y = 12;

            // 直接放置 CheckedListBox，不用 GroupBox
            clbShapeTypes = new CheckedListBox
            {
                Location = new Point(padding, y),
                Width = Width - padding * 2 - 12,
                Height = 170,
                CheckOnClick = true
            };
            DialogStyleHelper.StyleCheckedListBox(clbShapeTypes);

            foreach (var kvp in ShapeTypeMap)
            {
                clbShapeTypes.Items.Add(kvp.Key, true);
            }
            Controls.Add(clbShapeTypes);

            y += 178;

            // 全选 / 全不选
            btnSelectAll = new Button
            {
                Text = "全选",
                Width = 70,
                Height = 28,
                Location = new Point(padding, y)
            };
            DialogStyleHelper.StyleSecondaryButton(btnSelectAll);
            btnSelectAll.Click += (s, e) =>
            {
                for (int i = 0; i < clbShapeTypes.Items.Count; i++)
                    clbShapeTypes.SetItemChecked(i, true);
            };
            Controls.Add(btnSelectAll);

            btnUnselectAll = new Button
            {
                Text = "全不选",
                Width = 70,
                Height = 28,
                Location = new Point(padding + 78, y)
            };
            DialogStyleHelper.StyleSecondaryButton(btnUnselectAll);
            btnUnselectAll.Click += (s, e) =>
            {
                for (int i = 0; i < clbShapeTypes.Items.Count; i++)
                    clbShapeTypes.SetItemChecked(i, false);
            };
            Controls.Add(btnUnselectAll);

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
            btnOK.Click += (s, e) =>
            {
                if (clbShapeTypes.CheckedItems.Count == 0)
                {
                    MessageBox.Show("请至少选择一种图形类型。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };
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
    }
}
