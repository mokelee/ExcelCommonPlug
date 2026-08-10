using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 将工作表内容转换为 Markdown 表格格式。
    /// 第一行作为表头，其余行作为数据行。
    /// </summary>
    public class MarkdownService
    {
        private readonly Excel.Application _app;

        public MarkdownService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        /// <summary>
        /// 将指定工作表的已用区域转换为 Markdown 表格字符串。
        /// </summary>
        /// <param name="sheet">要转换的工作表</param>
        /// <returns>Markdown 格式的表格文本</returns>
        public string ConvertSheetToMarkdown(Excel.Worksheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            Excel.Range usedRange = sheet.UsedRange;
            if (usedRange == null)
                return string.Empty;

            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            if (rowCount == 0 || colCount == 0)
                return string.Empty;

            // 一次性读取所有单元格值到二维数组（性能远优于逐格读取）
            object[,] values;
            if (rowCount == 1 && colCount == 1)
            {
                // 单个单元格时 Value2 返回标量而非数组
                values = new object[2, 2];
                values[1, 1] = usedRange.Value2;
            }
            else
            {
                values = usedRange.Value2 as object[,];
                if (values == null)
                    return string.Empty;
            }

            var sb = new StringBuilder();

            // 第一行作为表头
            sb.Append('|');
            for (int col = 1; col <= colCount; col++)
            {
                sb.Append(' ');
                sb.Append(EscapeCell(values[1, col]));
                sb.Append(" |");
            }
            sb.AppendLine();

            // 分隔行
            sb.Append('|');
            for (int col = 1; col <= colCount; col++)
            {
                sb.Append(" --- |");
            }
            sb.AppendLine();

            // 数据行
            for (int row = 2; row <= rowCount; row++)
            {
                sb.Append('|');
                for (int col = 1; col <= colCount; col++)
                {
                    sb.Append(' ');
                    sb.Append(EscapeCell(values[row, col]));
                    sb.Append(" |");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        /// <summary>
        /// 转义单元格文本中对 Markdown 表格有影响的字符。
        /// </summary>
        private static string EscapeCell(object cellValue)
        {
            if (cellValue == null)
                return "";

            string text = cellValue.ToString();

            // 管道符是 Markdown 表格的列分隔符，必须转义
            text = text.Replace("|", "\\|");
            // 换行符替换为 <br>，保持表格结构完整
            text = text.Replace("\r\n", "<br>");
            text = text.Replace("\n", "<br>");
            text = text.Replace("\r", "<br>");

            return text;
        }
    }
}
