using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 区域解析结果：要么是可转换的有效区域，要么是需要提示用户的错误信息。
    /// </summary>
    public sealed class RangeResolveResult
    {
        /// <summary>解析成功时的有效内容区域（已裁掉边缘空行/空列）。</summary>
        public Excel.Range Range { get; }

        /// <summary>解析失败时给用户的提示文本；成功时为 null。</summary>
        public string ErrorMessage { get; }

        public bool Success => ErrorMessage == null;

        private RangeResolveResult(Excel.Range range, string errorMessage)
        {
            Range = range;
            ErrorMessage = errorMessage;
        }

        public static RangeResolveResult Ok(Excel.Range range) => new RangeResolveResult(range, null);
        public static RangeResolveResult Fail(string message) => new RangeResolveResult(null, message);
    }

    /// <summary>
    /// 将工作表内容转换为 Markdown 表格格式。
    /// 第一行作为表头，其余行作为数据行。
    /// </summary>
    public class MarkdownService
    {
        // 有效内容区域的行列上下限。
        private const int MaxColumns = 20;
        private const int MaxRows = 100;
        private const int MinRows = 2; // 至少 1 行表头 + 1 行数据

        private readonly Excel.Application _app;

        public MarkdownService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        /// <summary>
        /// 根据当前选区/整表确定要转换的有效区域，并校验其行列数是否在允许范围内。
        /// 规则：
        ///   - 用户选中了多个单元格 → 用选区；否则用整表 UsedRange。
        ///   - 取其中非空内容的最小外接矩形（裁掉边缘全空行/列）。
        ///   - 有效行数 &lt; 2 → 提示重选；列 &gt; 20 或行 &gt; 100 → 提示选更小区域。
        /// </summary>
        public RangeResolveResult ResolveConvertRange(Excel.Worksheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            // 判定用户是否主动选了一片区域（多于一个单元格）。
            Excel.Range candidate = null;
            var selection = _app.Selection as Excel.Range;
            if (selection != null && CountCells(selection) > 1)
                candidate = selection;
            else
                candidate = sheet.UsedRange;

            if (candidate == null)
                return RangeResolveResult.Fail("当前工作表没有数据。");

            // 收缩到非空内容的最小外接矩形。
            Excel.Range content = ShrinkToContent(candidate);
            if (content == null)
                return RangeResolveResult.Fail("所选区域没有可转换的内容，请重新选择。");

            int rows = content.Rows.Count;
            int cols = content.Columns.Count;

            if (rows < MinRows)
                return RangeResolveResult.Fail(
                    $"有效内容不足（至少需要 {MinRows} 行：1 行表头 + 1 行数据），请重新选择区域。");

            if (cols > MaxColumns || rows > MaxRows)
                return RangeResolveResult.Fail(
                    $"所选区域过大（有效内容 {rows} 行 × {cols} 列，上限 {MaxRows} 行 × {MaxColumns} 列），请手动选择更小的区域后再转换。");

            return RangeResolveResult.Ok(content);
        }

        /// <summary>
        /// 将指定区域转换为 Markdown 表格字符串。第一行作为表头。
        /// 直接读取每个单元格的 <c>Range.Text</c>（Excel 已渲染的显示文本），
        /// 保证日期、时间、百分比、货币、科学计数、分数、前导零、布尔、错误值等
        /// 与单元格实际显示完全一致，无需再自行套用数字格式。
        /// 额外处理超链接、合并单元格，以及列过窄导致的 "####" 兜底。
        /// </summary>
        public string ConvertRangeToMarkdown(Excel.Range range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;
            if (rowCount == 0 || colCount == 0)
                return string.Empty;

            var sb = new StringBuilder();

            // 表头
            AppendRow(sb, range, 1, colCount);

            // 分隔行
            sb.Append('|');
            for (int col = 1; col <= colCount; col++)
                sb.Append(" --- |");
            sb.AppendLine();

            // 数据行
            for (int row = 2; row <= rowCount; row++)
                AppendRow(sb, range, row, colCount);

            return sb.ToString();
        }

        private void AppendRow(StringBuilder sb, Excel.Range range, int row, int colCount)
        {
            sb.Append('|');
            for (int col = 1; col <= colCount; col++)
            {
                sb.Append(' ');
                sb.Append(EscapeCell(GetCellMarkdownText((Excel.Range)range.Cells[row, col])));
                sb.Append(" |");
            }
            sb.AppendLine();
        }

        /// <summary>
        /// 取单个单元格最终要写入 Markdown 的文本（转义前）。
        /// 合并单元格取左上角内容填充；有超链接则包装成 [文本](URL)。
        /// </summary>
        private static string GetCellMarkdownText(Excel.Range cell)
        {
            Excel.Range source = cell;

            // 合并单元格：非左上角格取左上角内容（Markdown 无合并概念，统一填充）。
            try
            {
                if (Convert.ToBoolean(cell.MergeCells))
                    source = (Excel.Range)cell.MergeArea.Cells[1, 1];
            }
            catch
            {
                // MergeCells 在混合区域可能返回 null，读取失败时按普通单元格处理。
            }

            string text = GetDisplayedCellText(source);
            return ConvertHyperlink(source, text);
        }

        /// <summary>
        /// 取单元格 Excel 渲染后的显示文本（Range.Text）。
        /// 当列过窄导致 Excel 显示为 "####" 时，回退到 Value2 的数值文本，
        /// 避免把占位符写进 Markdown。
        /// </summary>
        private static string GetDisplayedCellText(Excel.Range cell)
        {
            string text;
            try
            {
                text = cell.Text as string ?? cell.Text?.ToString() ?? "";
            }
            catch
            {
                text = "";
            }

            // "####" 兜底：列太窄时 Excel 用井号占位，回退到底层数值。
            if (text.Length > 0 && text.TrimEnd('#').Length == 0)
            {
                try
                {
                    object v = cell.Value2;
                    if (v is double d)
                        return FormatGeneralValue(d);
                    if (v != null)
                        return v.ToString();
                }
                catch
                {
                    // 回退失败时仍用原始的 "####"。
                }
            }

            return text;
        }

        /// <summary>
        /// 若单元格含超链接，包装为 Markdown 链接 [文本](URL)；否则原样返回文本。
        /// </summary>
        private static string ConvertHyperlink(Excel.Range cell, string text)
        {
            try
            {
                Excel.Hyperlinks links = cell.Hyperlinks;
                if (links == null || links.Count <= 0)
                    return text;

                Excel.Hyperlink link = links[1];
                string url = BuildHyperlinkUrl(link.Address, link.SubAddress);
                if (string.IsNullOrWhiteSpace(url))
                    return text;

                if (string.IsNullOrWhiteSpace(text))
                    text = url;

                // 链接文本内的方括号需转义，避免破坏 [文本] 结构。
                text = text.Replace("[", "\\[").Replace("]", "\\]");
                // URL 中的 ) 会提前闭合 Markdown 链接，编码掉。
                url = url.Replace(")", "%29");

                return $"[{text}]({url})";
            }
            catch
            {
                // 读取超链接失败时按普通文本输出。
                return text;
            }
        }

        /// <summary>
        /// 组合超链接 URL：外部地址、工作簿内部锚点（#Sheet!A1）或两者结合。
        /// </summary>
        private static string BuildHyperlinkUrl(string address, string subAddress)
        {
            bool hasAddress = !string.IsNullOrWhiteSpace(address);
            bool hasSubAddress = !string.IsNullOrWhiteSpace(subAddress);

            if (!hasAddress && !hasSubAddress) return null;
            if (hasAddress && !hasSubAddress) return address;
            if (!hasAddress && hasSubAddress) return "#" + subAddress;
            return $"{address}#{subAddress}";
        }

        /// <summary>
        /// 以接近 Excel 常规显示的方式输出数值：整数不带小数点，小数保留自然位数，
        /// 避免科学计数法。仅用于 "####" 兜底场景。
        /// </summary>
        private static string FormatGeneralValue(double d)
        {
            if (d == Math.Floor(d) && !double.IsInfinity(d) && Math.Abs(d) < 1e15)
                return ((long)d).ToString(System.Globalization.CultureInfo.InvariantCulture);

            return d.ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 转义单元格文本中对 Markdown 表格有影响的字符。
        /// 转义顺序：先反斜杠，再竖线，最后换行，避免二次转义相互干扰。
        /// 默认 Trim 首尾空白。
        /// </summary>
        private static string EscapeCell(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";

            // Trim 首尾空白（策略：不保留首尾空格）。
            text = text.Trim();
            if (text.Length == 0)
                return "";

            // 1) 反斜杠：先转义，避免后续新增的反斜杠被再次转义。
            text = text.Replace("\\", "\\\\");
            // 2) 管道符是 Markdown 表格的列分隔符，必须转义。
            text = text.Replace("|", "\\|");
            // 3) 美元符号：KaTeX/MathJax 渲染器会把 $...$ 当行内公式，
            //    含 $ 的文本（货币、会计格式代码等）需转义以避免被误解析。
            text = text.Replace("$", "\\$");
            // 4) 换行符替换为 <br>，保持表格结构完整。
            text = text.Replace("\r\n", "<br>");
            text = text.Replace("\n", "<br>");
            text = text.Replace("\r", "<br>");

            return text;
        }

        /// <summary>
        /// 将区域某属性读成 1-based 的二维数组（[rows+1, cols+1]，索引从 1 开始）。
        /// 单格区域返回标量，这里统一补成 2x2 数组便于按 [1,1] 访问。
        /// </summary>
        private static object[,] ReadAsMatrix(Excel.Range range, Func<Excel.Range, object> selector,
            int rowCount, int colCount)
        {
            if (rowCount == 1 && colCount == 1)
            {
                var single = new object[2, 2];
                single[1, 1] = selector(range);
                return single;
            }

            return selector(range) as object[,];
        }

        /// <summary>
        /// 统计区域单元格个数（区域可能是多块联合区域）。
        /// </summary>
        private static long CountCells(Excel.Range range)
        {
            try { return range.Count; }
            catch { return 0; }
        }

        /// <summary>
        /// 把区域收缩到「非空内容」的最小外接矩形：裁掉边缘的全空行与全空列。
        /// 若整块区域为空则返回 null。
        /// </summary>
        private static Excel.Range ShrinkToContent(Excel.Range range)
        {
            // 对联合区域/整行整列选择，先与已用区域求交集，避免遍历上百万空格。
            Excel.Worksheet sheet = range.Worksheet;
            Excel.Range used = sheet.UsedRange;
            Excel.Range area = _Intersect(range, used) ?? range;

            int rows = area.Rows.Count;
            int cols = area.Columns.Count;
            object[,] values = ReadAsMatrix(area, r => r.Value2, rows, cols);
            if (values == null)
                return null;

            int top = int.MaxValue, bottom = int.MinValue, left = int.MaxValue, right = int.MinValue;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    object v = values[r, c];
                    if (v == null) continue;
                    if (v is string s && s.Length == 0) continue;

                    if (r < top) top = r;
                    if (r > bottom) bottom = r;
                    if (c < left) left = c;
                    if (c > right) right = c;
                }
            }

            if (bottom < top || right < left)
                return null; // 全空

            Excel.Range topLeft = (Excel.Range)area.Cells[top, left];
            Excel.Range bottomRight = (Excel.Range)area.Cells[bottom, right];
            return sheet.Range[topLeft, bottomRight];
        }

        private static Excel.Range _Intersect(Excel.Range a, Excel.Range b)
        {
            if (a == null || b == null) return null;
            try { return a.Application.Intersect(a, b); }
            catch { return null; }
        }
    }
}
