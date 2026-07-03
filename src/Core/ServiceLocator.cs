using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCommonTools.Core
{
    /// <summary>
    /// 全局服务定位器，提供对 Excel Application 实例的静态访问。
    /// 在 ThisAddIn.Startup 中初始化。
    /// </summary>
    internal static class ServiceLocator
    {
        private static Excel.Application _application;

        public static void Initialize(Excel.Application application)
        {
            _application = application;
        }

        public static Excel.Application Application => _application;

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
