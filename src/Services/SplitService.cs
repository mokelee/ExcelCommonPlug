using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// Excel拆分服务：按列值拆分为多个工作表，或将各工作表分别保存为独立文件。
    /// </summary>
    public class SplitService
    {
        private readonly Excel.Application _app;

        public SplitService(Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 按指定列的值将数据拆分为多个工作表。
        /// </summary>
        /// <param name="headerRows">表头行数（从第1行开始）</param>
        /// <param name="splitColumn">拆分依据的列号（1-based）</param>
        /// <param name="saveAsFiles">拆分后是否将每个工作表另存为独立文件</param>
        /// <param name="includeEmpty">当拆分列存在空值时，是否将空值归入"NullEmpty"分组</param>
        /// <returns>拆分结果描述</returns>
        public string SplitByColumnToSheets(int headerRows, int splitColumn, bool saveAsFiles, bool includeEmpty = true)
        {
            if (headerRows < 1 || headerRows > 20)
                throw new ArgumentException("表头行数必须在1到20之间。");

            Excel.Worksheet sourceSheet = ServiceLocator.ActiveSheet;
            Excel.Workbook workbook = ServiceLocator.ActiveWorkbook;

            try
            {
                // 获取实际最后一行号（UsedRange.Row + UsedRange.Rows.Count - 1）
                int lastRow = sourceSheet.UsedRange.Row + sourceSheet.UsedRange.Rows.Count - 1;
                if (headerRows >= lastRow)
                    throw new InvalidOperationException("数据行数不足，表头行数大于等于总行数。");

                // 收集拆分列的唯一值
                var uniqueValues = new List<string>();
                var uniqueSet = new HashSet<string>(StringComparer.Ordinal);
                bool hasEmpty = false;

                for (int rIndex = headerRows + 1; rIndex <= lastRow; rIndex++)
                {
                    object cellValue = ((Excel.Range)sourceSheet.Cells[rIndex, splitColumn]).Value;
                    string textValue = cellValue != null ? Convert.ToString(cellValue) : "";

                    if (!string.IsNullOrEmpty(textValue))
                    {
                        if (uniqueSet.Add(textValue))
                        {
                            uniqueValues.Add(textValue);
                        }
                    }
                    else
                    {
                        hasEmpty = true;
                    }
                }

                if (hasEmpty)
                {
                    if (!includeEmpty)
                        throw new OperationCanceledException("拆分列存在空值，用户选择不包含空值。");

                    // 生成不与现有分类冲突的空值表名
                    string emptyName = "Empty";
                    int suffix = 1;
                    while (uniqueSet.Contains(emptyName))
                    {
                        emptyName = $"Empty_{suffix}";
                        suffix++;
                    }
                    uniqueSet.Add(emptyName);
                    uniqueValues.Add(emptyName);
                }

                // 确定空值对应的表名（后续分发时使用）
                string emptySheetName = null;
                if (hasEmpty && includeEmpty)
                {
                    emptySheetName = uniqueValues[uniqueValues.Count - 1];
                }

                // 验证所有名称是否合法且不重复
                foreach (string name in uniqueValues)
                {
                    if (!ValidateSheetName(name))
                        throw new InvalidOperationException($"名称 '{name}' 包含非法字符（< > / \\ | : \" * ?），请修改后重试。");

                    // 检查是否已存在同名工作表
                    try
                    {
                        Excel.Worksheet existing = (Excel.Worksheet)workbook.Sheets[name];
                        throw new InvalidOperationException($"当前工作簿中已存在名为 '{name}' 的工作表，无法自动拆分。");
                    }
                    catch (COMException)
                    {
                        // 工作表不存在，继续
                    }
                }

                using (new ExcelOperationScope(_app, suspendCalc: true))
                {
                    int sourceSheetIndex = sourceSheet.Index;

                    // 为每个唯一值创建新工作表并复制表头
                    foreach (string name in uniqueValues)
                    {
                        Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                        newSheet.Name = name;

                        // 复制表头行
                        Excel.Range headerRange = (Excel.Range)((Excel.Worksheet)workbook.Sheets[sourceSheetIndex]).Rows["1:" + headerRows];
                        headerRange.Copy();
                        ((Excel.Range)newSheet.Cells[1, 1]).PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);
                        headerRange.Copy(newSheet.Cells[1, 1]);
                        ComHelper.Release(headerRange);
                        ComHelper.Release(newSheet);
                    }

                    _app.CutCopyMode = (Excel.XlCutCopyMode)0;

                    // 将数据行分发到对应的工作表（使用字典跟踪每个目标表的下一行号）
                    var nextRowMap = new Dictionary<string, int>();
                    foreach (string name in uniqueValues)
                    {
                        nextRowMap[name] = headerRows + 1;
                    }

                    for (int rIndex = headerRows + 1; rIndex <= lastRow; rIndex++)
                    {
                        Excel.Range cellRange = (Excel.Range)sourceSheet.Cells[rIndex, splitColumn];
                        object cellValue = cellRange.Value;
                        ComHelper.Release(cellRange);

                        string textValue = cellValue != null ? Convert.ToString(cellValue) : "";
                        string sheetName = string.IsNullOrEmpty(textValue) ? emptySheetName : textValue;

                        Excel.Worksheet targetSheet = (Excel.Worksheet)workbook.Sheets[sheetName];
                        int targetNextRow = nextRowMap[sheetName];

                        ((Excel.Range)sourceSheet.Rows[rIndex]).Copy(targetSheet.Cells[targetNextRow, 1]);
                        nextRowMap[sheetName] = targetNextRow + 1;
                        ComHelper.Release(targetSheet);
                    }

                    _app.CutCopyMode = (Excel.XlCutCopyMode)0;

                    if (saveAsFiles)
                    {
                        SplitSheetsToFiles(uniqueValues);
                    }
                }

                return $"拆分完成，共创建 {uniqueValues.Count} 个工作表。";
            }
            finally
            {
                ComHelper.Release(sourceSheet);
                ComHelper.Release(workbook);
            }
        }

        /// <summary>
        /// 将当前工作簿中的每个工作表另存为独立的Excel文件。
        /// </summary>
        /// <param name="includeHidden">是否包含隐藏的工作表</param>
        public void SplitSheetsToFiles(bool includeHidden = false)
        {
            var allSheetNames = new List<string>();
            Excel.Workbook workbook = ServiceLocator.ActiveWorkbook;

            try
            {
                for (int i = 1; i <= workbook.Sheets.Count; i++)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[i];
                    // 跳过隐藏表（除非用户选择包含）
                    if (!includeHidden && sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    {
                        ComHelper.Release(sheet);
                        continue;
                    }
                    allSheetNames.Add(sheet.Name);
                    ComHelper.Release(sheet);
                }
            }
            finally
            {
                ComHelper.Release(workbook);
            }

            SplitSheetsToFiles(allSheetNames);
        }

        /// <summary>
        /// 将指定名称的工作表另存为独立文件。
        /// </summary>
        /// <param name="sheetNames">要保存的工作表名称列表</param>
        private void SplitSheetsToFiles(List<string> sheetNames)
        {
            Excel.Workbook workbook = ServiceLocator.ActiveWorkbook;
            string originalPath = workbook.Path;
            string originalName = workbook.Name;

            // 去掉扩展名
            int dotPos = originalName.LastIndexOf('.');
            string baseName = dotPos > 0 ? originalName.Substring(0, dotPos) : originalName;

            if (!originalPath.EndsWith("\\") && !originalPath.EndsWith("/"))
                originalPath += "\\";

            string activeWbName = workbook.Name;

            using (new ExcelOperationScope(_app, suspendAlerts: true))
            {
                try
                {
                    for (int index = 1; index <= workbook.Sheets.Count; index++)
                    {
                        Excel.Worksheet currentSheet = (Excel.Worksheet)workbook.Sheets[index];
                        string sheetName = currentSheet.Name;

                        // 检查是否在目标列表中
                        bool found = false;
                        foreach (string name in sheetNames)
                        {
                            if (name == sheetName)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            ComHelper.Release(currentSheet);
                            continue;
                        }

                        if (!ValidateSheetName(sheetName))
                        {
                            ComHelper.Release(currentSheet);
                            throw new InvalidOperationException($"工作表名 '{sheetName}' 包含非法字符，已跳过。");
                        }

                        // 如果工作表是隐藏的，先临时取消隐藏
                        Excel.XlSheetVisibility originalVisibility = currentSheet.Visible;
                        bool wasHidden = originalVisibility != Excel.XlSheetVisibility.xlSheetVisible;
                        if (wasHidden)
                        {
                            currentSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        }

                        // 复制工作表到新工作簿
                        currentSheet.Copy();

                        // 恢复隐藏状态
                        if (wasHidden)
                        {
                            currentSheet.Visible = originalVisibility;
                        }
                        ComHelper.Release(currentSheet);

                        Excel.Workbook newWb = _app.ActiveWorkbook;
                        string splitFileName = sheetName + ".xlsx";
                        string splitFilePath = originalPath + baseName + "_" + splitFileName;

                        // 断开外部链接
                        try
                        {
                            object[] links = (object[])newWb.LinkSources(Excel.XlLink.xlExcelLinks);
                            if (links != null)
                            {
                                foreach (object link in links)
                                {
                                    newWb.BreakLink((string)link, Excel.XlLinkType.xlLinkTypeExcelLinks);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[SplitService] 断开外部链接失败: {ex.Message}");
                        }

                        newWb.SaveAs(splitFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                        newWb.Close(false);

                        // 回到原始工作簿
                        _app.Workbooks[activeWbName].Activate();
                    }
                }
                finally
                {
                    ComHelper.Release(workbook);
                }
            }
        }

        /// <summary>
        /// 验证工作表名称是否合法（不包含Excel工作表名称中的非法字符）。
        /// </summary>
        /// <param name="name">待验证的名称</param>
        /// <returns>名称合法返回true，否则返回false</returns>
        public bool ValidateSheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return false;

            if (name.Length > 31)
                return false;

            char[] invalidChars = { '<', '>', '/', '\\', '|', ':', '"', '*', '?' };

            foreach (char c in invalidChars)
            {
                if (name.IndexOf(c) >= 0)
                    return false;
            }

            return true;
        }
    }
}
