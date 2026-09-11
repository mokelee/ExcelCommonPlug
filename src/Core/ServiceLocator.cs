using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCommonTools.Core
{
    /// <summary>
    /// 全局服务定位器，提供对 Excel Application 实例的静态访问。
    /// 在 AddIn.AutoOpen 中初始化，在 AutoClose 中调用 Reset 释放。
    /// </summary>
    internal static class ServiceLocator
    {
        private static volatile Excel.Application _application;

        public static void Initialize(Excel.Application application)
        {
            _application = application;
        }

        public static Excel.Application Application => _application;

        /// <summary>
        /// 释放对 Excel Application 的持有并置空引用。
        /// 在插件卸载（AutoClose）时调用，避免静态字段长期持有 RCW
        /// 导致 Excel 进程无法退出（后台残留幽灵 EXCEL.EXE）。
        /// </summary>
        public static void Reset()
        {
            var app = _application;
            _application = null;

            if (app != null)
            {
                try
                {
                    if (Marshal.IsComObject(app))
                    {
                        // 用 FinalReleaseComObject 一次性归零该 RCW 的引用计数。
                        Marshal.FinalReleaseComObject(app);
                    }
                }
                catch
                {
                    // 卸载阶段任何异常都不应抛出，静默吞掉。
                }
            }
        }

        /// <summary>
        /// 获取当前活动工作簿
        /// </summary>
        public static Excel.Workbook ActiveWorkbook => _application.ActiveWorkbook;

        public static Excel.Worksheet ActiveSheet
        {
            get
            {
                object sheet = _application.ActiveSheet;
                return sheet as Excel.Worksheet;
            }
        }
    }
}
