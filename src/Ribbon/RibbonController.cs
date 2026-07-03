using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCommonTools.Ribbon
{
    /// <summary>
    /// Ribbon 回调控制器（Excel-DNA 版本）。
    /// 继承 ExcelRibbon，处理 CustomRibbon.xml 中定义的所有按钮回调。
    /// </summary>
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private static Excel.Application App => Core.ServiceLocator.Application;
        private IRibbonUI _ribbonUI;
        private Services.SpotlightService _spotlightService;

        public override string GetCustomUI(string ribbonID)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string resourceName = "ExcelCommonTools.Ribbon.CustomRibbon.xml";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new FileNotFoundException($"找不到 Ribbon XML 资源: {resourceName}");
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
        }

        #region 文本处理

        public void OnTextBefore(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.InsertTextDialog("before"))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    Excel.Range selection = App.Selection as Excel.Range;
                    new Services.TextService(App).InsertTextBefore(selection, dialog.InputText);
                }
            }
            catch (Exception ex) { ShowError("前添加文本", ex); }
        }

        public void OnTextAfter(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.InsertTextDialog("after"))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    Excel.Range selection = App.Selection as Excel.Range;
                    new Services.TextService(App).InsertTextAfter(selection, dialog.InputText);
                }
            }
            catch (Exception ex) { ShowError("末尾添加文本", ex); }
        }

        public void OnTextMid(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.InsertTextDialog("mid"))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    Excel.Range selection = App.Selection as Excel.Range;
                    int position = dialog.MidPosition ?? 1;
                    new Services.TextService(App).InsertTextAtMiddle(selection, dialog.InputText, position, dialog.MidFromEnd);
                }
            }
            catch (Exception ex) { ShowError("中间插入文本", ex); }
        }

        #endregion

        #region 数据拆分

        public void OnSplitSheet(IRibbonControl control)
        {
            try
            {
                // 自动获取当前选中单元格的行和列
                Excel.Range selection = App.Selection as Excel.Range;
                int defaultHeaderRow = 1;
                string defaultCol = "A";

                if (selection != null)
                {
                    // 取左上角第一个单元格
                    Excel.Range firstCell = (Excel.Range)selection.Cells[1, 1];
                    defaultHeaderRow = Math.Max(1, firstCell.Row - 1);
                    defaultCol = ColumnNumberToLetter(firstCell.Column);
                }

                using (var dialog = new UI.SplitDialog(defaultHeaderRow, defaultCol))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    int splitColNum = ColumnLetterToNumber(dialog.SplitColumn);
                    if (splitColNum < 1)
                    {
                        MessageBox.Show("拆分列字母无效。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    string result = new Services.SplitService(App).SplitByColumnToSheets(dialog.HeaderRows, splitColNum, dialog.SaveAsFiles);
                    MessageBox.Show(result, "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { ShowError("按列拆分成Sheet", ex); }
        }

        public void OnSplitFile(IRibbonControl control)
        {
            try
            {
                // 弹出选项对话框
                using (var dlg = new Form())
                {
                    dlg.Text = "拆分表为独立文件";
                    dlg.Width = 320;
                    dlg.Height = 150;
                    dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                    dlg.MaximizeBox = false;
                    dlg.MinimizeBox = false;
                    dlg.StartPosition = FormStartPosition.CenterScreen;

                    var chk = new CheckBox
                    {
                        Text = "包含隐藏的工作表",
                        Location = new System.Drawing.Point(20, 20),
                        AutoSize = true,
                        Checked = false
                    };
                    dlg.Controls.Add(chk);

                    var btnOK = new Button
                    {
                        Text = "确定",
                        DialogResult = DialogResult.OK,
                        Width = 80,
                        Height = 30,
                        Location = new System.Drawing.Point(60, 65)
                    };
                    dlg.Controls.Add(btnOK);

                    var btnCancel = new Button
                    {
                        Text = "取消",
                        DialogResult = DialogResult.Cancel,
                        Width = 80,
                        Height = 30,
                        Location = new System.Drawing.Point(160, 65)
                    };
                    dlg.Controls.Add(btnCancel);

                    dlg.AcceptButton = btnOK;
                    dlg.CancelButton = btnCancel;

                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    new Services.SplitService(App).SplitSheetsToFiles(chk.Checked);
                }
                MessageBox.Show("拆分为独立文件完成。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { ShowError("拆分为独立文件", ex); }
        }

        #endregion

        #region 图片处理

        public void OnResizeImg(IRibbonControl control)
        {
            try
            {
                new Services.ImageService(App).ResizeImagesToCells();
                MessageBox.Show("调整图片大小完成。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { ShowError("调整图片大小", ex); }
        }

        public void OnInsertImg(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.InsertImageDialog(1, 2))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    new Services.ImageService(App).InsertImagesFromFolder(
                        dialog.LeftCol, dialog.RightCol, dialog.StartRow, dialog.FolderPath, dialog.PicExtension);
                    MessageBox.Show("批量插入图片完成。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { ShowError("批量插入图片", ex); }
        }

        public void OnExportImg(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.ExportImageDialog(1))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    new Services.ImageService(App).ExportImagesToFile(dialog.NameCol, dialog.FolderPath, dialog.PicExtension);
                    MessageBox.Show($"导出所有图片完成，文件已保存到：\n{dialog.FolderPath}",
                        "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { ShowError("导出所有图片", ex); }
        }

        #endregion

        #region 工作表

        public void OnUnhide(IRibbonControl control)
        {
            try
            {
                var sheetService = new Services.SheetService(App);
                var hiddenSheets = sheetService.GetHiddenSheets();
                if (hiddenSheets.Count == 0)
                {
                    MessageBox.Show("当前工作簿没有隐藏的工作表。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var hiddenNames = hiddenSheets.Select(s => s.Name).ToList();
                using (var dialog = new UI.UnhideSheetDialog(hiddenNames))
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    var selectedIndices = new List<int>();
                    foreach (string name in dialog.UnhideSheetNames)
                    {
                        var matched = hiddenSheets.FirstOrDefault(s => s.Name == name);
                        if (!string.IsNullOrEmpty(matched.Name))
                            selectedIndices.Add(matched.Index);
                    }
                    if (selectedIndices.Count > 0)
                    {
                        sheetService.UnhideSheets(selectedIndices);
                        MessageBox.Show("取消隐藏完成。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { ShowError("取消隐藏", ex); }
        }

        #endregion

        #region 图形/批注

        public void OnDelShape(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.DeleteShapeDialog())
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    int count = new Services.ShapeService(App).DeleteShapesByTypes(Core.ServiceLocator.ActiveSheet, dialog.SelectedShapeTypes);
                    MessageBox.Show($"删除完成，共删除 {count} 个图形。\n\n提示：删除批注后的红色三角标记为临时显示，保存并重新打开文件后会消失。",
                        "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { ShowError("删除指定图形", ex); }
        }

        public void OnComment(IRibbonControl control)
        {
            try
            {
                using (var dialog = new UI.CommentDialog())
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    bool? showAll = dialog.VisibilityMode == "show" ? true : (bool?)null;
                    bool? hideAll = dialog.VisibilityMode == "hide" ? true : (bool?)null;
                    int count = new Services.CommentService(App).AdjustComments(
                        Core.ServiceLocator.ActiveSheet,
                        dialog.IsRight, dialog.HorizontalGap,
                        dialog.IsAbove, dialog.VerticalGap,
                        dialog.CommentWidth, dialog.CommentHeight,
                        showAll, hideAll);
                    MessageBox.Show($"调整批注完成，共处理 {count} 个批注。", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { ShowError("调整批注", ex); }
        }

        #endregion

        #region 聚光灯

        public void OnSpotlightToggle(IRibbonControl control, bool pressed)
        {
            try
            {
                if (_spotlightService == null)
                {
                    _spotlightService = new Services.SpotlightService(App);
                }

                if (pressed)
                {
                    _spotlightService.Enable();
                }
                else
                {
                    _spotlightService.Disable();
                }
            }
            catch (Exception ex) { ShowError("聚光灯", ex); }
        }

        public bool GetSpotlightPressed(IRibbonControl control)
        {
            return _spotlightService?.IsEnabled ?? false;
        }

        public Bitmap GetSpotlightImage(IRibbonControl control)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream("ExcelCommonTools.Resources.spotlight.png"))
            {
                if (stream == null) return null;
                var bmp = new Bitmap(stream);
                // 将白色/近白色背景替换为透明
                bmp.MakeTransparent(Color.White);
                return bmp;
            }
        }

        #endregion

        #region 关于

        public string GetVersionLabel(IRibbonControl control)
        {
            return $"v{Core.AppConstants.AppVersion}";
        }

        public void OnAbout(IRibbonControl control)
        {
            MessageBox.Show(
                $"{Core.AppConstants.AppName}\n\n" +
                $"版本：{Core.AppConstants.AppVersion}\n" +
                $"产品ID：{Core.AppConstants.ProductId}\n\n" +
                $"Excel 日常办公辅助工具插件",
                "关于",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        #endregion

        #region 辅助

        private static void ShowError(string operation, Exception ex)
        {
            MessageBox.Show($"操作【{operation}】失败：\n{ex.Message}", "日常工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static int ColumnLetterToNumber(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter)) return -1;
            columnLetter = columnLetter.Trim().ToUpper();
            int number = 0;
            foreach (char c in columnLetter)
            {
                if (c < 'A' || c > 'Z') return -1;
                number = number * 26 + (c - 'A' + 1);
            }
            return number;
        }

        private static string ColumnNumberToLetter(int colNum)
        {
            string result = "";
            while (colNum > 0)
            {
                colNum--;
                result = (char)('A' + colNum % 26) + result;
                colNum /= 26;
            }
            return result;
        }

        #endregion
    }
}
