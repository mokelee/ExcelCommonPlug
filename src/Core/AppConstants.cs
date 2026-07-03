using System.Reflection;

namespace ExcelCommonTools.Core
{
    /// <summary>
    /// 应用程序常量
    /// </summary>
    internal static class AppConstants
    {
        // 应用信息
        public const string AppName = "Excel日常工具";
        public const string ProductId = "excel-common-tools";

        // 版本号：从程序集版本读取，编译时由 csproj 的 Version 属性决定
        public static readonly string AppVersion = typeof(AppConstants).Assembly.GetName().Version.ToString(3);

        // 升级服务器配置
        public const string ServerUrl = "http://10.18.98.146:8100";
    }
}
