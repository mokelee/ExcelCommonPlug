using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 区域解析结果：要么是可转换的「逻辑行列表」，要么是需要提示用户的错误信息。
    /// 逻辑行列表已完成隐藏行/列过滤、多块选区校验与取值（Excel 显示文本 + 转义前处理）。
    /// </summary>
    public sealed class RangeResolveResult
    {
        /// <summary>解析成功时的行列表；每行是若干个「转义前」的单元格文本。成功时非 null。</summary>
        public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

        /// <summary>解析失败时给用户的提示文本；成功时为 null。</summary>
        public string ErrorMessage { get; }

        public bool Success => ErrorMessage == null;

        private RangeResolveResult(IReadOnlyList<IReadOnlyList<string>> rows, string errorMessage)
        {
            Rows = rows;
            ErrorMessage = errorMessage;
        }

        public static RangeResolveResult Ok(IReadOnlyList<IReadOnlyList<string>> rows)
            => new RangeResolveResult(rows, null);
        public static RangeResolveResult Fail(string message)
            => new RangeResolveResult(null, message);
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
        /// 根据当前选区/整表确定要转换的逻辑行列表，并校验规模。
        /// 规则：
        ///   - 用户主动选了区域（&gt;1 格）→ 用选区（可能是多块联合区域 Areas）；否则用整表 UsedRange。
        ///   - 多块选区：要求各块的物理列区间完全相同，否则报错；随后按块顺序纵向拼接所有行。
        ///   - 跳过隐藏行与隐藏列（筛选隐藏、手动隐藏均算）。
        ///   - 全空行保留（用户既然选了就输出），仅裁掉整表模式下的边缘空白。
        ///   - 有效行数 &lt; 2 → 提示重选；列 &gt; 20 或行 &gt; 100 → 提示选更小区域。
        /// </summary>
        public RangeResolveResult ResolveConvertRange(Excel.Worksheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            // 判定用户是否主动选了区域（多于一个单元格）。
            var selection = _app.Selection as Excel.Range;
            bool userSelected = selection != null && CountCells(selection) > 1;

            Excel.Range candidate = userSelected ? selection : sheet.UsedRange;
            if (candidate == null)
                return RangeResolveResult.Fail("当前工作表没有数据。");

            // 收集各块的物理列（跳过隐藏列），并做多块一致性校验。
            List<Excel.Range> areas = GetAreas(candidate);
            if (areas.Count == 0)
                return RangeResolveResult.Fail("所选区域没有可转换的内容，请重新选择。");

            // 多块选区：物理列区间必须完全相同。
            if (areas.Count > 1)
            {
                string columnSignature = null;
                foreach (Excel.Range area in areas)
                {
                    string sig = GetColumnSignature(area);
                    if (columnSignature == null)
                        columnSignature = sig;
                    else if (!string.Equals(columnSignature, sig, StringComparison.Ordinal))
                        return RangeResolveResult.Fail(
                            "所选各区域的列不一致（列位置或列数不同），无法转换。请选择列完全相同的区域。");
                }
            }

            // 计算参与输出的可见物理列（以第一块为准，多块已校验列相同）。
            List<int> visibleColumns = GetVisibleColumns(sheet, areas[0]);
            if (visibleColumns.Count == 0)
                return RangeResolveResult.Fail("所选区域的列全部被隐藏，没有可转换的内容。");

            // 按块顺序纵向收集可见行。整表模式下裁掉边缘空白行；多块选区保留全空行。
            var rows = new List<IReadOnlyList<string>>();
            foreach (Excel.Range area in areas)
                CollectRows(sheet, area, visibleColumns, rows);

            // 整表模式（单块且来自 UsedRange）裁掉首尾的全空行，避免 UsedRange 带出的空白。
            if (!userSelected)
                TrimEdgeEmptyRows(rows);

            if (rows.Count < MinRows)
                return RangeResolveResult.Fail(
                    $"有效内容不足（至少需要 {MinRows} 行：1 行表头 + 1 行数据），请重新选择区域。");

            int cols = visibleColumns.Count;
            if (cols > MaxColumns || rows.Count > MaxRows)
                return RangeResolveResult.Fail(
                    $"所选区域过大（有效内容 {rows.Count} 行 × {cols} 列，上限 {MaxRows} 行 × {MaxColumns} 列），请手动选择更小的区域后再转换。");

            return RangeResolveResult.Ok(rows);
        }

        /// <summary>
        /// 将已解析的逻辑行列表渲染为 Markdown 表格字符串。第一行作为表头。
        /// </summary>
        public string ConvertRangeToMarkdown(RangeResolveResult resolved)
        {
            if (resolved == null)
                throw new ArgumentNullException(nameof(resolved));
            if (!resolved.Success || resolved.Rows == null || resolved.Rows.Count == 0)
                return string.Empty;

            IReadOnlyList<IReadOnlyList<string>> rows = resolved.Rows;
            int colCount = rows[0].Count;

            var sb = new StringBuilder();

            // 表头
            AppendRow(sb, rows[0], colCount);

            // 分隔行
            sb.Append('|');
            for (int c = 0; c < colCount; c++)
                sb.Append(" --- |");
            sb.AppendLine();

            // 数据行
            for (int r = 1; r < rows.Count; r++)
                AppendRow(sb, rows[r], colCount);

            return sb.ToString();
        }

        private static void AppendRow(StringBuilder sb, IReadOnlyList<string> cells, int colCount)
        {
            sb.Append('|');
            for (int c = 0; c < colCount; c++)
            {
                string raw = c < cells.Count ? cells[c] : "";
                sb.Append(' ');
                sb.Append(EscapeCell(raw));
                sb.Append(" |");
            }
            sb.AppendLine();
        }

        /// <summary>
        /// 展开联合区域为各块（Areas）；单块区域返回自身。空区域返回空列表。
        /// </summary>
        private static List<Excel.Range> GetAreas(Excel.Range range)
        {
            var result = new List<Excel.Range>();
            if (range == null)
                return result;

            try
            {
                Excel.Areas areas = range.Areas;
                for (int i = 1; i <= areas.Count; i++)
                    result.Add(areas[i]);
            }
            catch
            {
                result.Add(range);
            }

            return result;
        }

        /// <summary>
        /// 生成某块的物理列签名「起始列:列数」，用于多块「列位置与列数完全相同」校验。
        /// 连续列区间下，起始列与列数即可唯一确定列范围。
        /// </summary>
        private static string GetColumnSignature(Excel.Range area)
        {
            int startCol = area.Column;
            int colCount = area.Columns.Count;
            return startCol + ":" + colCount;
        }

        /// <summary>
        /// 取某块中可见（未隐藏）的物理列号列表，从左到右。
        /// </summary>
        private static List<int> GetVisibleColumns(Excel.Worksheet sheet, Excel.Range area)
        {
            var cols = new List<int>();
            int startCol = area.Column;
            int colCount = area.Columns.Count;

            for (int i = 0; i < colCount; i++)
            {
                int physicalCol = startCol + i;
                if (!IsColumnHidden(sheet, physicalCol))
                    cols.Add(physicalCol);
            }
            return cols;
        }

        /// <summary>
        /// 收集某块中可见行的取值（跳过隐藏行），追加到 rows。
        /// 每行按 visibleColumns 指定的物理列取值。
        /// </summary>
        private static void CollectRows(Excel.Worksheet sheet, Excel.Range area,
            List<int> visibleColumns, List<IReadOnlyList<string>> rows)
        {
            int startRow = area.Row;
            int rowCount = area.Rows.Count;

            for (int i = 0; i < rowCount; i++)
            {
                int physicalRow = startRow + i;
                if (IsRowHidden(sheet, physicalRow))
                    continue;

                var cells = new List<string>(visibleColumns.Count);
                foreach (int col in visibleColumns)
                    cells.Add(GetCellMarkdownText((Excel.Range)sheet.Cells[physicalRow, col]));

                rows.Add(cells);
            }
        }

        /// <summary>
        /// 裁掉行列表首尾的「全空」行（仅用于整表 UsedRange 模式）。
        /// </summary>
        private static void TrimEdgeEmptyRows(List<IReadOnlyList<string>> rows)
        {
            // 尾部
            while (rows.Count > 0 && IsRowEmpty(rows[rows.Count - 1]))
                rows.RemoveAt(rows.Count - 1);
            // 头部
            while (rows.Count > 0 && IsRowEmpty(rows[0]))
                rows.RemoveAt(0);
        }

        private static bool IsRowEmpty(IReadOnlyList<string> row)
        {
            foreach (string cell in row)
                if (!string.IsNullOrEmpty(cell) && cell.Trim().Length > 0)
                    return false;
            return true;
        }

        private static bool IsRowHidden(Excel.Worksheet sheet, int row)
        {
            try { return Convert.ToBoolean(((Excel.Range)sheet.Rows[row]).Hidden); }
            catch { return false; }
        }

        private static bool IsColumnHidden(Excel.Worksheet sheet, int column)
        {
            try { return Convert.ToBoolean(((Excel.Range)sheet.Columns[column]).Hidden); }
            catch { return false; }
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
                return ((long)d).ToString(CultureInfo.InvariantCulture);

            return d.ToString("0.###############", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 转义单元格文本中对 Markdown 表格有影响的字符。
        /// 转义顺序：先反斜杠，再竖线，再美元号，最后换行，避免二次转义相互干扰。
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
        /// 统计区域单元格个数（区域可能是多块联合区域）。
        /// </summary>
        private static long CountCells(Excel.Range range)
        {
            try { return range.Count; }
            catch { return 0; }
        }
    }
}
