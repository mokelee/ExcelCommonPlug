using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelCommonTools.UI
{
    public class UnhideSheetDialog : Form
    {
        private readonly List<string> _hiddenSheetNames;
        private readonly List<string> _unhideSheetNames = new List<string>();

        private ListBox lstHideList;
        private ListBox lstUnHideList;
        private Button btnAdd;
        private Button btnDel;
        private Button btnOK;
        private Button btnCancel;

        public List<int> SelectedSheetIndices { get; private set; } = new List<int>();

        public UnhideSheetDialog(List<string> hiddenSheetNames)
        {
            _hiddenSheetNames = hiddenSheetNames ?? throw new ArgumentNullException(nameof(hiddenSheetNames));

            Text = "取消隐藏工作表";
            Width = 550;
            Height = 420;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            DialogStyleHelper.ApplyFormStyle(this);
            InitializeComponent();

            // 填充隐藏工作表列表
            foreach (var name in _hiddenSheetNames)
            {
                lstHideList.Items.Add(name);
            }
        }

        private void InitializeComponent()
        {
            int padding = 20;
            int listWidth = 165;
            int listHeight = 260;
            int midBtnWidth = 140;
            int midBtnX = padding + listWidth + 10;
            int rightListX = midBtnX + midBtnWidth + 10;
            int y = 16;

            // 左侧标签
            var lblHide = new Label
            {
                Text = "隐藏的工作表：",
                Location = new Point(padding, y),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblHide);
            Controls.Add(lblHide);

            // 右侧标签
            var lblUnHide = new Label
            {
                Text = "取消隐藏：",
                Location = new Point(rightListX, y),
                AutoSize = true
            };
            DialogStyleHelper.StyleLabel(lblUnHide);
            Controls.Add(lblUnHide);

            y += 24;

            // 左侧列表
            lstHideList = new ListBox
            {
                Location = new Point(padding, y),
                Width = listWidth,
                Height = listHeight,
                SelectionMode = SelectionMode.MultiExtended
            };
            DialogStyleHelper.StyleListBox(lstHideList);
            Controls.Add(lstHideList);

            // 中间按钮（垂直居中）
            int midBtnY = y + (listHeight - 160) / 2;

            btnAdd = new Button
            {
                Text = "添加 >>",
                Width = midBtnWidth,
                Height = 32,
                Location = new Point(midBtnX, midBtnY)
            };
            DialogStyleHelper.StylePrimaryButton(btnAdd);
            btnAdd.Click += BtnAdd_Click;
            Controls.Add(btnAdd);

            var btnAddAll = new Button
            {
                Text = "全部添加 >>",
                Width = midBtnWidth,
                Height = 32,
                Location = new Point(midBtnX, midBtnY + 40)
            };
            DialogStyleHelper.StyleSecondaryButton(btnAddAll);
            btnAddAll.Click += (s, ev) =>
            {
                while (lstHideList.Items.Count > 0)
                {
                    string name = lstHideList.Items[0].ToString();
                    lstHideList.Items.RemoveAt(0);
                    lstUnHideList.Items.Add(name);
                    _unhideSheetNames.Add(name);
                }
            };
            Controls.Add(btnAddAll);

            btnDel = new Button
            {
                Text = "<< 移除",
                Width = midBtnWidth,
                Height = 32,
                Location = new Point(midBtnX, midBtnY + 88)
            };
            DialogStyleHelper.StyleSecondaryButton(btnDel);
            btnDel.Click += BtnDel_Click;
            Controls.Add(btnDel);

            var btnDelAll = new Button
            {
                Text = "<< 全部移除",
                Width = midBtnWidth,
                Height = 32,
                Location = new Point(midBtnX, midBtnY + 128)
            };
            DialogStyleHelper.StyleSecondaryButton(btnDelAll);
            btnDelAll.Click += (s, ev) =>
            {
                while (lstUnHideList.Items.Count > 0)
                {
                    string name = lstUnHideList.Items[0].ToString();
                    lstUnHideList.Items.RemoveAt(0);
                    lstHideList.Items.Add(name);
                    _unhideSheetNames.Remove(name);
                }
            };
            Controls.Add(btnDelAll);

            // 右侧列表
            lstUnHideList = new ListBox
            {
                Location = new Point(rightListX, y),
                Width = listWidth,
                Height = listHeight,
                SelectionMode = SelectionMode.MultiExtended
            };
            DialogStyleHelper.StyleListBox(lstUnHideList);
            Controls.Add(lstUnHideList);

            // 底部按钮栏
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

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            var toMove = new List<string>();
            foreach (var item in lstHideList.SelectedItems)
                toMove.Add(item.ToString());

            foreach (var name in toMove)
            {
                lstHideList.Items.Remove(name);
                lstUnHideList.Items.Add(name);
                _unhideSheetNames.Add(name);
            }
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            var toMove = new List<string>();
            foreach (var item in lstUnHideList.SelectedItems)
                toMove.Add(item.ToString());

            foreach (var name in toMove)
            {
                lstUnHideList.Items.Remove(name);
                lstHideList.Items.Add(name);
                _unhideSheetNames.Remove(name);
            }
        }

        /// <summary>
        /// 要取消隐藏的工作表名称列表
        /// </summary>
        public List<string> UnhideSheetNames => _unhideSheetNames.ToList();
    }
}
