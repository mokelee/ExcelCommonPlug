using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 工作表服务：隐藏/取消隐藏工作表、清理空行空表、移除工作表保护等。
    /// </summary>
    public class SheetService
    {
        private readonly Excel.Application _app;

        public SheetService(Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 获取当前工作簿中所有隐藏的工作表。
        /// </summary>
        /// <returns>隐藏工作表的名称和索引列表</returns>
        public List<(string Name, int Index)> GetHiddenSheets()
        {
            var result = new List<(string Name, int Index)>();
            Excel.Workbook workbook = ServiceLocator.ActiveWorkbook;

            for (int index = 1; index <= workbook.Sheets.Count; index++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[index];
                // xlSheetVisible = -1, xlSheetHidden = 0, xlSheetVeryHidden = 2
                if ((int)sheet.Visible != -1)
                {
                    result.Add((sheet.Name, sheet.Index));
                }
            }

            return result;
        }

        /// <summary>
        /// 取消隐藏指定的工作表。
        /// </summary>
        /// <param name="sheetIndices">要取消隐藏的工作表索引列表（1-based）</param>
        public void UnhideSheets(List<int> sheetIndices)
        {
            if (sheetIndices == null || sheetIndices.Count == 0)
                return;

            Excel.Workbook workbook = ServiceLocator.ActiveWorkbook;

            foreach (int index in sheetIndices)
            {
                if (index >= 1 && index <= workbook.Sheets.Count)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[index];
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                }
            }
        }

        /// <summary>
        /// 清理指定工作表中所有完全为空的行（从底部向上删除）。
        /// </summary>
        /// <param name="sheet">要清理的工作表</param>
        public void CleanEmptyRows(Excel.Worksheet sheet)
        {
            if (sheet == null) return;

            using (new ExcelOperationScope(_app, suspendCalc: true))
            {
                int lastRow = sheet.UsedRange.Row - 1 + sheet.UsedRange.Rows.Count;

                for (int nowRow = lastRow; nowRow >= 1; nowRow--)
                {
                    Excel.Range rowRange = (Excel.Range)sheet.Rows[nowRow];
                    if (_app.WorksheetFunction.CountA(rowRange) == 0)
                    {
                        rowRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    }
                    ComHelper.Release(rowRange);
                }
            }
        }

        /// <summary>
        /// 清理所有工作表：删除空白工作表、清除空白行和空白列、删除隐藏的图形对象。
        /// </summary>
        public void CleanEmptySheetsAndShapes()
        {
            Excel.Workbook workbook = ServiceLocator.ActiveWorkbook;

            using (new ExcelOperationScope(_app, suspendAlerts: true))
            {
                // 从后往前遍历，避免删除时索引变化
                for (int sIdx = workbook.Worksheets.Count; sIdx >= 1; sIdx--)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets[sIdx];

                    if (_app.WorksheetFunction.CountA(sheet.Cells) == 0)
                    {
                        // 工作表完全为空，删除
                        // 至少保留一个工作表
                        if (workbook.Worksheets.Count > 1)
                        {
                            sheet.Delete();
                        }
                    }
                    else
                    {
                        // 找到最后有数据的行和列
                        Excel.Range lastDataRow = sheet.Cells.Find(
                            "*",
                            sheet.Cells[1, 1],
                            Excel.XlFindLookIn.xlValues,
                            Excel.XlLookAt.xlWhole,
                            Excel.XlSearchOrder.xlByRows,
                            Excel.XlSearchDirection.xlPrevious,
                            false
                        );

                        Excel.Range lastDataCol = sheet.Cells.Find(
                            "*",
                            sheet.Cells[1, 1],
                            Excel.XlFindLookIn.xlValues,
                            Excel.XlLookAt.xlWhole,
                            Excel.XlSearchOrder.xlByColumns,
                            Excel.XlSearchDirection.xlPrevious,
                            false
                        );

                        if (lastDataRow != null && lastDataCol != null)
                        {
                            int lastRow = lastDataRow.Row;
                            int lastCol = lastDataCol.Column;
                            int totalRows = sheet.Rows.Count;
                            int totalCols = sheet.Columns.Count;

                            // 清除并删除最后数据行以下的所有行
                            if (lastRow < totalRows)
                            {
                                Excel.Range rowsToDelete = sheet.get_Range(
                                    sheet.Cells[lastRow + 1, 1],
                                    sheet.Cells[totalRows, 1]
                                );
                                rowsToDelete.EntireRow.Clear();
                                rowsToDelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                                ComHelper.Release(rowsToDelete);
                            }

                            // 删除最后数据列以右的所有列
                            if (lastCol < totalCols)
                            {
                                Excel.Range colsToDelete = sheet.get_Range(
                                    sheet.Cells[1, lastCol + 1],
                                    sheet.Cells[1, totalCols]
                                );
                                colsToDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                                ComHelper.Release(colsToDelete);
                            }
                        }

                        ComHelper.Release(lastDataRow);
                        ComHelper.Release(lastDataCol);

                        // 删除隐藏的图形对象
                        foreach (Excel.Shape shape in sheet.Shapes)
                        {
                            try
                            {
                                if ((int)shape.Visible == 0) // msoFalse = 0
                                {
                                    shape.Delete();
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[SheetService] 删除图形对象失败: {ex.Message}");
                            }
                            finally
                            {
                                ComHelper.Release(shape);
                            }
                        }
                    }

                    ComHelper.Release(sheet);
                }
            }
        }

        /// <summary>
        /// 移除指定工作表的保护。
        /// 通过反复切换保护状态来解除工作表保护。
        /// </summary>
        /// <param name="sheet">要解除保护的工作表</param>
        public void RemoveSheetProtection(Excel.Worksheet sheet)
        {
            if (sheet == null) return;

            // 通过反复切换保护选项来解除保护
            sheet.Protect(Type.Missing, true, true, Type.Missing, Type.Missing,
                true, true, true, true, true, true, true, true, true, true, true);
            sheet.Protect(Type.Missing, false, true, Type.Missing, Type.Missing,
                true, true, true, true, true, true, true, true, true, true, true);
            sheet.Protect(Type.Missing, true, true, Type.Missing, Type.Missing,
                true, true, true, true, true, true, true, true, true, true, true);
            sheet.Protect(Type.Missing, false, true, Type.Missing, Type.Missing,
                true, true, true, true, true, true, true, true, true, true, true);
        }
    }
}
