using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 批注服务：批量调整工作表中批注的位置、大小和显示状态。
    /// </summary>
    public class CommentService
    {
        private readonly Excel.Application _app;

        public CommentService(Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 批量调整工作表中所有批注的位置、大小和显示状态。
        /// </summary>
        /// <param name="sheet">要操作的工作表</param>
        /// <param name="isRight">true=批注在单元格右侧，false=左侧</param>
        /// <param name="hGap">水平间距（磅）：批注框边缘到单元格边缘的距离</param>
        /// <param name="isAbove">true=批注在单元格上方，false=下方</param>
        /// <param name="vGap">垂直间距（磅）：批注框边缘到单元格边缘的距离</param>
        /// <param name="width">批注宽度，null表示不修改</param>
        /// <param name="height">批注高度，null表示不修改</param>
        /// <param name="showAll">true=显示所有批注，null表示不修改</param>
        /// <param name="hideAll">true=隐藏所有批注，null表示不修改</param>
        /// <returns>处理的批注数量</returns>
        public int AdjustComments(Excel.Worksheet sheet, bool isRight, double hGap,
            bool isAbove, double vGap, double? width, double? height,
            bool? showAll, bool? hideAll)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            int count = 0;

            using (new ExcelOperationScope(_app))
            {
                foreach (Excel.Comment comment in sheet.Comments)
                {
                    count++;
                    Excel.Range parentCell = null;
                    Excel.Shape commentShape = null;

                    try
                    {
                        parentCell = (Excel.Range)comment.Parent;
                        commentShape = comment.Shape;

                        double cellLeft = (double)parentCell.Left;
                        double cellTop = (double)parentCell.Top;
                        double cellWidth = (double)parentCell.Width;
                        double cellHeight = (double)parentCell.Height;
                        float shapeWidth = commentShape.Width;
                        float shapeHeight = commentShape.Height;

                        // 调整宽度（先设置，因为后续位置计算可能依赖宽高）
                        if (width.HasValue)
                        {
                            commentShape.Width = (float)width.Value;
                            shapeWidth = (float)width.Value;
                        }

                        // 调整高度
                        if (height.HasValue)
                        {
                            commentShape.Height = (float)height.Value;
                            shapeHeight = (float)height.Value;
                        }

                        // 水平位置
                        if (isRight)
                        {
                            // 右侧：批注左边缘 = 单元格右边缘 + 间距
                            commentShape.Left = (float)(cellLeft + cellWidth + hGap);
                        }
                        else
                        {
                            // 左侧：批注右边缘 = 单元格左边缘 - 间距
                            commentShape.Left = (float)(cellLeft - shapeWidth - hGap);
                        }

                        // 垂直位置
                        if (isAbove)
                        {
                            // 上方：批注顶部 = 单元格顶部 - 间距
                            commentShape.Top = (float)(cellTop - vGap);
                        }
                        else
                        {
                            // 下方：批注顶部 = 单元格底部 + 间距
                            commentShape.Top = (float)(cellTop + cellHeight + vGap);
                        }

                        // 显示或隐藏批注
                        if (showAll.HasValue && showAll.Value)
                        {
                            commentShape.Visible = Office.MsoTriState.msoTrue;
                        }
                        else if (hideAll.HasValue && hideAll.Value)
                        {
                            commentShape.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[CommentService] 调整批注失败: {ex.Message}");
                    }
                    finally
                    {
                        ComHelper.Release(parentCell);
                        ComHelper.Release(commentShape);
                    }
                }
            }

            return count;
        }
    }
}
