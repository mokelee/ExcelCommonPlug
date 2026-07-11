using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelCommonTools.Core;

namespace ExcelCommonTools
{
    /// <summary>
    /// Excel-DNA 加载项入口点。
    /// 负责初始化全局服务，并在后台通过轻量 API 检查更新。
    /// 如有更新，启动独立的 UpdateClient.exe 执行升级。
    /// </summary>
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
#if DEBUG
            string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "update_check.log");
#endif
            try
            {
#if DEBUG
                File.AppendAllText(logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoOpen started. Version={AppConstants.AppVersion}\r\n");
#endif

                var app = (Excel.Application)ExcelDnaUtil.Application;
                ServiceLocator.Initialize(app);

                // 初始化日志系统（安装路径下 logs 目录）
                try
                {
                    string xllDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
                    string logsDir = Path.Combine(xllDir, "logs");
                    Logger.Init(logsDir);
                    Logger.Info("AddIn", $"AutoOpen started. Version={AppConstants.AppVersion}, XllPath={ExcelDnaUtil.XllPath}");
                }
                catch (Exception logEx)
                {
                    Debug.WriteLine($"[AddIn] Logger.Init failed: {logEx.Message}");
#if DEBUG
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Logger.Init failed: {logEx.Message}\r\n{logEx.StackTrace}\r\n");
#endif
                }

                // 注册未处理异常捕获（用于排查 Spotlight 等功能导致 Excel 崩溃）
                AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
                Application.ThreadException += OnThreadException;

                // 后台检查更新（不阻塞 Excel 启动）
                Task.Run(() => CheckForUpdates());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddIn] AutoOpen 初始化失败: {ex.Message}");
                Logger.Error("AddIn", $"AutoOpen 初始化失败: {ex.Message}", ex);
#if DEBUG
                try { File.AppendAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "update_check.log"), $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoOpen 失败: {ex.Message}\r\n{ex.StackTrace}\r\n\r\n"); } catch { }
#endif
            }
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            Logger.Error("UnhandledException", ex?.Message ?? "Unknown fatal error", ex);
        }

        private static void OnThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            Logger.Error("ThreadException", e.Exception?.Message ?? "Unknown thread error", e.Exception);
        }

        public void AutoClose()
        {
        }

        /// <summary>
        /// 轻量级版本检查：调用 /check?version=x.y.z API。
        /// 如有更新，在 UI 线程提示用户，确认后启动 UpdateClient.exe 并退出 Excel。
        /// </summary>
        private async void CheckForUpdates()
        {
            try
            {
                string installDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);

                // 读取本地版本
                string localVersion = AppConstants.AppVersion;
                string versionFile = Path.Combine(installDir, ".version");
                if (File.Exists(versionFile))
                {
                    string v = File.ReadAllText(versionFile).Trim();
                    if (!string.IsNullOrEmpty(v)) localVersion = v;
                }

                // 轻量级版本检查
                string checkUrl = $"{AppConstants.ServerUrl}/api/v1/products/{AppConstants.ProductId}/check?version={localVersion}";

                using (var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) })
                {
                    var response = await http.GetAsync(checkUrl).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) return;

                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    Logger.Debug("UpdateCheck", $"URL: {checkUrl}, Response: {json}");

                    var result = ParseCheckResult(json);

                    if (result == null || !result.HasUpdate) return;

                    // 在 UI 线程提示用户
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        string forceText = result.ForceUpdate ? "\n此为强制更新，必须升级后才能使用。" : "";
                        string notesText = !string.IsNullOrEmpty(result.ReleaseNotes) ? $"\n\n更新内容：{result.ReleaseNotes}" : "";
                        var dlgResult = MessageBox.Show(
                            $"发现新版本 {result.LatestVersion}，当前版本 {localVersion}。{forceText}{notesText}\n\n是否立即更新？",
                            AppConstants.AppName,
                            MessageBoxButtons.YesNo,
                            result.ForceUpdate ? MessageBoxIcon.Exclamation : MessageBoxIcon.Information);

                        if (dlgResult == DialogResult.Yes)
                        {
                            // 二次确认：提醒用户保存文件
                            var confirmResult = MessageBox.Show(
                                "升级将关闭 Excel，请确保已保存所有文件。\n\n点击\"确定\"开始升级。",
                                AppConstants.AppName,
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Warning);

                            if (confirmResult == DialogResult.OK)
                            {
                                LaunchUpdateClient(installDir);
                            }
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddIn] 版本检查失败: {ex.Message}");
                Logger.Error("UpdateCheck", $"版本检查失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 启动独立的 UpdateClient.exe 并退出 Excel。
        /// UpdateClient 会等待 Excel 进程退出后执行文件替换。
        /// </summary>
        private void LaunchUpdateClient(string installDir)
        {
            try
            {
                string updaterPath = Path.Combine(installDir, "UpdateClient.exe");
                if (!File.Exists(updaterPath))
                {
                    MessageBox.Show("未找到升级程序 UpdateClient.exe，请联系管理员。",
                        AppConstants.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 构建命令行参数
                string args = $"--server-url \"{AppConstants.ServerUrl}\" " +
                              $"--product-id \"{AppConstants.ProductId}\" " +
                              $"--install-dir \"{installDir}\" " +
                              $"--wait-process EXCEL " +
                              $"--restart-command excel.exe";

                Process.Start(new ProcessStartInfo
                {
                    FileName = updaterPath,
                    Arguments = args,
                    UseShellExecute = true
                });

                // 退出 Excel，让 UpdateClient 可以替换文件
                ServiceLocator.Application.Quit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddIn] 启动 UpdateClient 失败: {ex.Message}");
                MessageBox.Show($"启动升级程序失败：{ex.Message}",
                    AppConstants.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 版本检查 API 响应模型
        /// </summary>
        private class CheckResult
        {
            public bool HasUpdate { get; set; }
            public string LatestVersion { get; set; }
            public bool ForceUpdate { get; set; }
            public string DownloadUrl { get; set; }
            public string ReleaseNotes { get; set; }
        }

        /// <summary>
        /// 简单 JSON 解析（避免 System.Text.Json 在 LoadFromBytes 模式下的兼容性问题）
        /// </summary>
        private static CheckResult ParseCheckResult(string json)
        {
            if (string.IsNullOrEmpty(json)) return null;
            var result = new CheckResult();
            result.HasUpdate = GetJsonBool(json, "hasUpdate");
            result.ForceUpdate = GetJsonBool(json, "forceUpdate");
            result.LatestVersion = GetJsonString(json, "latestVersion");
            result.DownloadUrl = GetJsonString(json, "downloadUrl");
            result.ReleaseNotes = GetJsonString(json, "releaseNotes");
            return result;
        }

        private static string GetJsonString(string json, string key)
        {
            string pattern = $"\"{key}\"";
            int idx = json.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return null;
            int colonIdx = json.IndexOf(':', idx + pattern.Length);
            if (colonIdx < 0) return null;
            int startQuote = json.IndexOf('"', colonIdx + 1);
            if (startQuote < 0) return null;
            int endQuote = json.IndexOf('"', startQuote + 1);
            if (endQuote < 0) return null;
            return json.Substring(startQuote + 1, endQuote - startQuote - 1);
        }

        private static bool GetJsonBool(string json, string key)
        {
            string pattern = $"\"{key}\"";
            int idx = json.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return false;
            int colonIdx = json.IndexOf(':', idx + pattern.Length);
            if (colonIdx < 0) return false;
            string rest = json.Substring(colonIdx + 1, Math.Min(10, json.Length - colonIdx - 1)).Trim();
            return rest.StartsWith("true", StringComparison.OrdinalIgnoreCase);
        }
    }
}
