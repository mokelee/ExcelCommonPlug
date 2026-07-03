using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCommonTools.Core
{
    /// <summary>
    /// Excel 操作作用域，自动管理 ScreenUpdating、Calculation、DisplayAlerts 的保存与恢复。
    /// 使用 using 语句确保操作完成后自动恢复 Excel 的原始状态。
    /// 
    /// 用法示例：
    /// <code>
    /// using (new ExcelOperationScope(app))
    /// {
    ///     // 业务逻辑，ScreenUpdating 已自动关闭
    /// }
    /// // 退出 using 时自动恢复 ScreenUpdating
    /// 
    /// using (new ExcelOperationScope(app, suspendCalc: true, suspendAlerts: true))
    /// {
    ///     // ScreenUpdating、Calculation、DisplayAlerts 均已暂停
    /// }
    /// </code>
    /// </summary>
    public class ExcelOperationScope : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly bool _oldScreenUpdating;
        private readonly Excel.XlCalculation _oldCalculation;
        private readonly bool _oldDisplayAlerts;
        private readonly bool _suspendCalc;
        private readonly bool _suspendAlerts;
        private bool _disposed;

        /// <summary>
        /// 创建操作作用域，立即关闭 ScreenUpdating 并保存原始状态。
        /// </summary>
        /// <param name="app">Excel Application 实例</param>
        /// <param name="suspendCalc">是否同时暂停自动计算</param>
        /// <param name="suspendAlerts">是否同时关闭警告提示</param>
        public ExcelOperationScope(Excel.Application app, bool suspendCalc = false, bool suspendAlerts = false)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _suspendCalc = suspendCalc;
            _suspendAlerts = suspendAlerts;

            // 保存原始状态
            _oldScreenUpdating = _app.ScreenUpdating;
            _oldCalculation = _app.Calculation;
            _oldDisplayAlerts = _app.DisplayAlerts;

            // 应用优化设置
            _app.ScreenUpdating = false;
            if (_suspendCalc)
                _app.Calculation = Excel.XlCalculation.xlCalculationManual;
            if (_suspendAlerts)
                _app.DisplayAlerts = false;
        }

        /// <summary>
        /// 恢复 Excel 的原始状态。
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try { _app.ScreenUpdating = _oldScreenUpdating; } catch { }
            if (_suspendCalc)
            {
                try { _app.Calculation = _oldCalculation; } catch { }
            }
            if (_suspendAlerts)
            {
                try { _app.DisplayAlerts = _oldDisplayAlerts; } catch { }
            }
        }
    }
}
