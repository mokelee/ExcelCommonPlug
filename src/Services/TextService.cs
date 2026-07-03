using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 文本处理服务，提供文本前后插入、中间插入、大写转换、提取汉字/数字/字母、删除空格/字符等功能。
    /// </summary>
    public class TextService
    {
        private readonly Excel.Application _app;

        /// <summary>
        /// 创建文本处理服务实例
        /// </summary>
        /// <param name="app">Excel Application 实例</param>
        public TextService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        /// <summary>
        /// 在选中单元格的文本前添加指定文本
        /// </summary>
        /// <param name="selection">选中的单元格区域</param>
        /// <param name="insertText">要添加的文本内容</param>
        public void InsertTextBefore(Excel.Range selection, string insertText)
        {
            if (selection == null || string.IsNullOrEmpty(insertText)) return;

            using (new ExcelOperationScope(_app))
            {
                foreach (Excel.Range cell in selection)
                {
                    object val = cell.Value2;
                    if (val != null)
                    {
                        cell.Value2 = insertText + val.ToString();
                    }
                    ComHelper.Release(cell);
                }
            }
        }

        /// <summary>
        /// 在选中单元格的文本末尾追加指定文本
        /// </summary>
        /// <param name="selection">选中的单元格区域</param>
        /// <param name="insertText">要追加的文本内容</param>
        public void InsertTextAfter(Excel.Range selection, string insertText)
        {
            if (selection == null || string.IsNullOrEmpty(insertText)) return;

            using (new ExcelOperationScope(_app))
            {
                foreach (Excel.Range cell in selection)
                {
                    object val = cell.Value2;
                    if (val != null)
                    {
                        cell.Value2 = val.ToString() + insertText;
                    }
                    ComHelper.Release(cell);
                }
            }
        }

        /// <summary>
        /// 在选中单元格的文本中间指定位置插入文本。
        /// 当 fromEnd 为 false 时，从开头第 position 个字符后插入；
        /// 当 fromEnd 为 true 时，从末尾第 position 个字符前插入。
        /// </summary>
        /// <param name="selection">选中的单元格区域</param>
        /// <param name="insertText">要插入的文本内容</param>
        /// <param name="position">插入位置（1-based）</param>
        /// <param name="fromEnd">是否从末尾开始计算位置</param>
        public void InsertTextAtMiddle(Excel.Range selection, string insertText, int position, bool fromEnd = false)
        {
            if (selection == null || string.IsNullOrEmpty(insertText)) return;

            using (new ExcelOperationScope(_app))
            {
                foreach (Excel.Range cell in selection)
                {
                    string cellText = (string)cell.Text;
                    if (string.IsNullOrEmpty(cellText))
                    {
                        ComHelper.Release(cell);
                        continue;
                    }

                    int textLen = cellText.Length;
                    int actualPos = fromEnd ? textLen - position : position;

                    if (actualPos <= 0)
                    {
                        cell.Value2 = insertText + cellText;
                    }
                    else if (actualPos >= textLen)
                    {
                        cell.Value2 = cellText + insertText;
                    }
                    else
                    {
                        string left = cellText.Substring(0, actualPos);
                        string right = cellText.Substring(actualPos);
                        cell.Value2 = left + insertText + right;
                    }
                    ComHelper.Release(cell);
                }
            }
        }

        /// <summary>
        /// 将指定单元格的数值转换为中文大写货币格式。
        /// 使用 Excel 的 TEXT 函数配合 [DBNUM2] 格式实现转换。
        /// 例如：123.45 → 壹佰贰拾叁元肆角伍分整
        /// </summary>
        /// <param name="cell">目标单元格（将写入公式到其右侧相邻单元格）</param>
        public void ConvertToChineseCurrency(Excel.Range cell)
        {
            if (cell == null) return;

            Excel.Range targetCell = null;
            try
            {
                // 在目标单元格右侧写入中文大写转换公式
                // =TEXT(INT(源单元格),"[DBNUM2]")&"元"&TEXT(MID(源单元格,LEN(INT(源单元格))+2,1),"[DBNUM2]D角")
                // &TEXT(MID(源单元格,LEN(INT(源单元格))+3,1),"[DBNUM2]D分")&"整"
                targetCell = cell.Offset[0, 1] as Excel.Range;
                string srcAddr = cell.Address[false, false, Excel.XlReferenceStyle.xlR1C1];

                string formula =
                    $"=TEXT(INT({srcAddr}),\"[DBNUM2]\")&\"元\"" +
                    $"&TEXT(MID({srcAddr},LEN(INT({srcAddr}))+2,1),\"[DBNUM2]D角\")" +
                    $"&TEXT(MID({srcAddr},LEN(INT({srcAddr}))+3,1),\"[DBNUM2]D分\")" +
                    $"&\"整\"";

                targetCell.FormulaR1C1 = formula;
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"[TextService] ConvertToChineseCurrency: {ex.Message}");
            }
            finally
            {
                ComHelper.Release(targetCell);
            }
        }

    }
}
