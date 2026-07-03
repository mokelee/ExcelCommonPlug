using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 图形对象服务：获取图形类型、按类型删除图形。
    /// </summary>
    public class ShapeService
    {
        private readonly Excel.Application _app;

        public ShapeService(Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 获取可用的图形类型及其对应的整数值。
        /// </summary>
        /// <returns>图形类型名称到整数值的映射字典</returns>
        public Dictionary<string, int> GetShapeTypes()
        {
            return new Dictionary<string, int>
            {
                { "批注", (int)Office.MsoShapeType.msoComment },       // msoComment = 4
                { "图片", (int)Office.MsoShapeType.msoPicture },       // msoPicture = 13
                { "文本框", (int)Office.MsoShapeType.msoTextBox },     // msoTextBox = 17
                { "SmartArt", (int)Office.MsoShapeType.msoSmartArt },  // msoSmartArt = 15
                { "图表", (int)Office.MsoShapeType.msoChart },         // msoChart = 3
                { "形状", (int)Office.MsoShapeType.msoAutoShape }      // msoAutoShape = 1
            };
        }

        /// <summary>
        /// 删除指定工作表中匹配指定类型的图形对象。
        /// 批注通过 Comments 集合单独处理。
        /// SmartArt 通过 HasSmartArt 属性辅助判断。
        /// </summary>
        /// <param name="sheet">要操作的工作表</param>
        /// <param name="shapeTypes">要删除的图形类型值列表</param>
        /// <returns>删除的图形数量</returns>
        public int DeleteShapesByTypes(Excel.Worksheet sheet, List<int> shapeTypes)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));
            if (shapeTypes == null || shapeTypes.Count == 0)
                return 0;

            int deleteCount = 0;
            bool deleteComments = shapeTypes.Contains(4); // msoComment = 4
            bool deleteSmartArt = shapeTypes.Contains(25); // msoSmartArt = 25

            using (new ExcelOperationScope(_app))
            {
                // 遍历图形对象，从后往前删除
                for (int i = sheet.Shapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = null;
                    try
                    {
                        shape = sheet.Shapes.Item(i);
                        int shapeType = (int)shape.Type;
                        bool shouldDelete = false;

                        // 精确类型匹配
                        foreach (int targetType in shapeTypes)
                        {
                            if (shapeType == targetType)
                            {
                                shouldDelete = true;
                                break;
                            }
                        }

                        // SmartArt 特殊处理：SmartArt 的 Shape.Type 可能是 msoGroup(6) 或 msoSmartArt
                        // 通过 HasSmartArt 属性辅助判断
                        if (!shouldDelete && deleteSmartArt)
                        {
                            try
                            {
                                if (shape.HasSmartArt == Office.MsoTriState.msoTrue)
                                    shouldDelete = true;
                            }
                            catch { }
                        }

                        if (shouldDelete)
                        {
                            shape.Delete();
                            deleteCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ShapeService] 删除图形对象失败: {ex.Message}");
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }

                // 批注/备注：清除所有批注和备注（包括红色三角标记）
                if (deleteComments)
                {
                    try
                    {
                        Excel.Range usedRange = sheet.UsedRange;
                        try
                        {
                            // SpecialCells(xlCellTypeComments) 定位所有有批注/备注的单元格
                            Excel.Range commentCells = usedRange.SpecialCells(
                                Excel.XlCellType.xlCellTypeComments);
                            if (commentCells != null)
                            {
                                foreach (Excel.Range cell in commentCells)
                                {
                                    try
                                    {
                                        // 清除备注（红色三角，旧版 Comment/Note）
                                        if (cell.Comment != null)
                                        {
                                            cell.Comment.Delete();
                                        }
                                        // 同时尝试清除 NoteText
                                        try { cell.NoteText(""); } catch { }
                                        deleteCount++;
                                    }
                                    catch { }
                                    finally
                                    {
                                        ComHelper.Release(cell);
                                    }
                                }
                                ComHelper.Release(commentCells);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // 没有批注/备注单元格
                        }
                        ComHelper.Release(usedRange);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ShapeService] 删除批注/备注失败: {ex.Message}");
                    }
                }
            }

            return deleteCount;
        }
    }
}
