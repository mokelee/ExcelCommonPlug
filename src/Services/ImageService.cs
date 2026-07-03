using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// Excel图片服务：调整图片大小、从文件夹插入图片、导出图片到文件。
    /// </summary>
    public class ImageService
    {
        private readonly Excel.Application _app;

        public ImageService(Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 将活动工作表中所有图片调整为所在单元格大小。
        /// 支持合并单元格：使用MergeArea获取合并区域的宽高。
        /// 保持宽高比，居中显示。
        /// </summary>
        public void ResizeImagesToCells()
        {
            Excel.Worksheet sheet = ServiceLocator.ActiveSheet;

            using (new ExcelOperationScope(_app))
            {
                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    if ((int)shape.Type != (int)Office.MsoShapeType.msoPicture) continue;

                    // 获取图片所在单元格（支持合并单元格）
                    Excel.Range targetRange = shape.TopLeftCell.MergeArea;

                    double cellWidth = (double)targetRange.Width;
                    double cellHeight = (double)targetRange.Height;
                    double cellTop = (double)targetRange.Top;
                    double cellLeft = (double)targetRange.Left;

                    // 计算边距（根据单元格大小动态调整）
                    int deltaHeight = (int)(cellHeight / 100 + 4);
                    int deltaWidth = (int)(cellWidth / 100 + 4);

                    // 锁定宽高比，先恢复原始大小
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                    shape.ScaleHeight((float)1, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                    shape.Placement = Excel.XlPlacement.xlMoveAndSize;

                    // 计算缩放比例
                    double scaleByHeight = (cellHeight - deltaHeight) / shape.Height;
                    double scaleByWidth = (cellWidth - deltaWidth) / shape.Width;

                    // 设置图片位置（先放到单元格左上角 + 边距）
                    shape.Top = (float)(cellTop + deltaHeight / 2.0);
                    shape.Left = (float)(cellLeft + deltaWidth / 2.0);

                    // 根据较小的缩放比例决定以宽还是高为基准
                    if (scaleByHeight < scaleByWidth)
                    {
                        // 以高度为基准
                        shape.Height = (float)(cellHeight - deltaHeight);
                    }
                    else
                    {
                        // 以宽度为基准
                        shape.Width = (float)(cellWidth - deltaWidth);
                        // 垂直居中
                        shape.Top = (float)(cellTop + (cellHeight - shape.Height) / 2.0 - 1);
                    }

                    ComHelper.Release(targetRange);
                }
            }
        }

        /// <summary>
        /// 从指定文件夹批量插入图片到工作表。
        /// 根据左列单元格的值查找对应图片文件，插入到右列单元格中并缩放适配。
        /// </summary>
        /// <param name="leftCol">名称列号（1-based），单元格值用于匹配图片文件名</param>
        /// <param name="rightCol">图片插入列号（1-based）</param>
        /// <param name="startRow">起始行号</param>
        /// <param name="folderPath">图片文件夹路径</param>
        /// <param name="picExtension">图片扩展名（如 ".png"、".jpg"）</param>
        public void InsertImagesFromFolder(int leftCol, int rightCol, int startRow, string folderPath, string picExtension)
        {
            if (string.IsNullOrEmpty(folderPath))
                throw new ArgumentException("文件夹路径不能为空。");

            Excel.Worksheet sheet = ServiceLocator.ActiveSheet;

            using (new ExcelOperationScope(_app))
            {
                // 先删除工作表中已有的图片
                for (int i = sheet.Shapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = sheet.Shapes.Item(i);
                    if ((int)shape.Type == (int)Office.MsoShapeType.msoPicture)
                    {
                        shape.Delete();
                    }
                    ComHelper.Release(shape);
                }

                // 获取最后一行
                int lastRow = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1;
                string excelPath = ServiceLocator.ActiveWorkbook.Path;

                for (int rIdx = startRow; rIdx <= lastRow; rIdx++)
                {
                    object cellValue = ((Excel.Range)sheet.Cells[rIdx, leftCol]).Value;
                    if (cellValue == null || string.IsNullOrEmpty(Convert.ToString(cellValue)))
                        continue;

                    string picName = Convert.ToString(cellValue);

                    // 构建图片路径：优先使用指定文件夹，如果文件夹为空则使用工作簿路径下的"产品图片"子目录
                    string picFullPath;
                    if (!string.IsNullOrEmpty(folderPath))
                    {
                        picFullPath = Path.Combine(folderPath, picName + picExtension);
                    }
                    else
                    {
                        picFullPath = Path.Combine(excelPath, "产品图片", picName + picExtension);
                    }

                    if (!File.Exists(picFullPath))
                        continue;

                    Excel.Range targetCell = (Excel.Range)sheet.Cells[rIdx, rightCol];

                    // 插入图片（初始尺寸设为较大值，后续缩放）
                    Excel.Shape newShape = sheet.Shapes.AddPicture(
                        picFullPath,
                        Office.MsoTriState.msoFalse,   // LinkToFile
                        Office.MsoTriState.msoTrue,    // SaveWithDocument
                        (float)(Convert.ToDouble(targetCell.Left) + 2),
                        (float)(Convert.ToDouble(targetCell.Top) + 2),
                        300,        // 初始宽度
                        600         // 初始高度
                    );

                    // 锁定宽高比，恢复原始大小
                    newShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                    newShape.ScaleHeight((float)1, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                    newShape.Placement = Excel.XlPlacement.xlMoveAndSize;

                    double cellHeight = Convert.ToDouble(targetCell.Height);
                    double cellWidth = Convert.ToDouble(targetCell.Width);
                    double cellTop = Convert.ToDouble(targetCell.Top);
                    double cellLeft = Convert.ToDouble(targetCell.Left);

                    // 计算缩放比例
                    double scaleByHeight = (cellHeight - 4) / newShape.Height;
                    double scaleByWidth = (cellWidth - 4) / newShape.Width;

                    // 设置位置
                    newShape.Top = (float)(cellTop + 2);
                    newShape.Left = (float)(cellLeft + 2);

                    // 根据较小的缩放比例调整
                    if (scaleByHeight < scaleByWidth)
                    {
                        newShape.Height = (float)(cellHeight - 4);
                    }
                    else
                    {
                        newShape.Width = (float)(cellWidth - 4);
                    }

                    // 垂直居中
                    double topDelta = (cellHeight - newShape.Height) / 2.0 - 1;
                    if (topDelta > 2)
                    {
                        newShape.Top = (float)(cellTop + topDelta);
                    }

                    ComHelper.Release(newShape);
                    ComHelper.Release(targetCell);
                }
            }
        }

        /// <summary>
        /// 将活动工作表中的所有图片导出为文件。
        /// 通过匹配图片Top位置与指定列的单元格来确定文件名。
        /// 使用Chart中转方式导出图片。
        /// </summary>
        /// <param name="nameCol">名称列号（1-based），单元格值用于命名导出文件</param>
        /// <param name="outputFolder">导出目标文件夹路径</param>
        /// <param name="picExtension">图片扩展名（如 ".png"、".jpg"）</param>
        public void ExportImagesToFile(int nameCol, string outputFolder, string picExtension)
        {
            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentException("导出文件夹路径不能为空。");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            Excel.Worksheet sheet = ServiceLocator.ActiveSheet;

            using (new ExcelOperationScope(_app))
            {
                int lastRow = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count;

                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    if ((int)shape.Type != (int)Office.MsoShapeType.msoPicture) continue;

                    // 根据图片的Top位置匹配对应的行
                    string picName = "";
                    double shapeTop = shape.Top;

                    for (int rIdx = 1; rIdx <= lastRow; rIdx++)
                    {
                        Excel.Range cell = (Excel.Range)sheet.Cells[rIdx, nameCol];
                        double cellTop = Convert.ToDouble(cell.Top);
                        double cellHeight = Convert.ToDouble(cell.Height);

                        if (shapeTop > cellTop && shapeTop < cellTop + cellHeight)
                        {
                            object cellValue = cell.Value;
                            if (cellValue != null)
                            {
                                picName = Convert.ToString(cellValue) + picExtension;
                            }
                            ComHelper.Release(cell);
                            break;
                        }

                        ComHelper.Release(cell);
                    }

                    if (string.IsNullOrEmpty(picName)) continue;

                    string outputPath = Path.Combine(outputFolder, picName);

                    // 使用Chart中转方式导出图片
                    Excel.ChartObject chartObj = null;
                    Excel.Chart chart = null;
                    try
                    {
                        // 短暂延迟确保选中完成
                        Thread.Sleep(200);

                        shape.Copy();

                        // 创建临时图表对象
                        Excel.ChartObjects chartObjects = (Excel.ChartObjects)sheet.ChartObjects();
                        chartObj = chartObjects.Add(1, 1, shape.Width - 1, shape.Height - 1);
                        ComHelper.Release(chartObjects);
                        chart = chartObj.Chart;

                        chartObj.Activate();
                        chartObj.Border.LineStyle = (Excel.XlLineStyle)0;

                        Thread.Sleep(100);

                        // 粘贴图片到图表
                        chart.Paste();

                        Thread.Sleep(200);

                        // 导出图片（Filtername取扩展名去掉点号的后3个字符）
                        string filterName = picExtension.TrimStart('.');
                        chart.Export(outputPath, filterName, false);

                        Thread.Sleep(200);

                        // 删除临时图表
                        chartObj.Delete();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ImageService] 导出图片失败: {ex.Message}");
                    }
                    finally
                    {
                        ComHelper.Release(chart);
                        ComHelper.Release(chartObj);
                    }
                }
            }
        }
    }
}
