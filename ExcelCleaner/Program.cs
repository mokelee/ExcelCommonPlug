using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelCleaner
{
    /// <summary>
    /// 清理后台残留的 Excel 幽灵进程。
    ///
    /// 背景：在装有 AppLocker/SRP 的锁定环境（如企业域）里，脚本无法启动 powershell.exe，
    /// 因此原先的「powershell + clean_excel.ps1」清理方案会被拦截、无法运行。
    /// 本程序是独立编译的 Native AOT 单文件 exe，零运行时依赖，可在锁定环境正常运行。
    ///
    /// 策略 A：只结束「没有主窗口」的 EXCEL.EXE（MainWindowHandle == 0）。
    ///   - 正在使用、有可见窗口的 Excel 会被保留，不会误关、不丢数据。
    ///   - 后台残留（无主窗口）的幽灵进程会被清理，恢复插件正常加载。
    ///
    /// 用法：
    ///   ExcelCleaner.exe            手动模式：清理后弹窗反馈结果。
    ///   ExcelCleaner.exe --silent   静默模式：不弹窗，用退出码反馈（供安装程序调用）。
    ///
    /// 静默模式退出码：
    ///   0  - 已清理干净：没有残留的后台进程，且没有正在使用的 Excel 窗口。
    ///   10 - 仍有正在使用的 Excel 窗口（有主窗口），未关闭；调用方应提示用户手动关闭。
    ///   11 - 有幽灵进程因权限不足无法结束（拒绝访问）；调用方可提示以管理员身份重试。
    /// </summary>
    internal static class Program
    {
        // ---- user32 MessageBox（原生，AOT 友好，无托管 UI 依赖） ----
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        // ---- 窗口枚举（用于判断进程是否拥有「可见的顶层窗口」，比 Process.MainWindowHandle 更稳）----
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLengthW(IntPtr hWnd);

        private const uint MB_OK = 0x0;
        private const uint MB_ICONINFORMATION = 0x40;
        private const uint MB_ICONWARNING = 0x30;

        private const string Title = "Excel日常工具 - 清理后台进程";

        // 退出码
        private const int ExitClean = 0;
        private const int ExitVisibleRemains = 10;
        private const int ExitAccessDenied = 11;

        private static int Main(string[] args)
        {
            bool silent = false;
            foreach (var a in args)
            {
                if (string.Equals(a, "--silent", StringComparison.OrdinalIgnoreCase))
                    silent = true;
            }

            var ghosts = new List<Process>();
            var visible = new List<Process>();
            ClassifyExcelProcesses(ghosts, visible);

            // 没有幽灵进程
            if (ghosts.Count == 0)
            {
                if (!silent)
                {
                    if (visible.Count > 0)
                        ShowInfo($"未发现后台残留的 Excel 进程。{VisibleNote(visible.Count)}");
                    else
                        ShowInfo("未发现任何 Excel 进程，无需清理。");
                }
                return visible.Count > 0 ? ExitVisibleRemains : ExitClean;
            }

            // 记录第一次分类得到的可见 Excel 数量，用于退出码/提示。
            // 不做清理后的二次枚举：那会有时序竞争（进程刚被结束时窗口句柄可能瞬时读不到），
            // 而「有窗口的正在使用的 Excel」本来就不在清理目标里，数量不会因清理而变化。
            int visibleCount = visible.Count;

            int killed = 0;
            var deniedIds = new List<int>();
            var otherErrors = new List<string>();

            foreach (var p in ghosts)
            {
                int procId;
                try { procId = p.Id; } catch { continue; }

                bool gone = TryKill(p, out string? denyReason);
                if (gone)
                {
                    killed++;
                }
                else if (denyReason == "denied")
                {
                    deniedIds.Add(procId);
                }
                else
                {
                    otherErrors.Add($"PID {procId}: {denyReason}");
                }
            }

            if (silent)
            {
                if (deniedIds.Count > 0) return ExitAccessDenied;
                if (visibleCount > 0) return ExitVisibleRemains;
                return ExitClean;
            }

            // 手动模式：汇总反馈
            var lines = new List<string>();
            if (killed > 0)
            {
                lines.Add($"已清理 {killed} 个后台残留的 Excel 进程。");
                lines.Add("现在可以重新打开 Excel。");
            }
            if (deniedIds.Count > 0)
            {
                lines.Add($"仍有 {deniedIds.Count} 个进程因权限不足无法清理（PID: {string.Join(", ", deniedIds)}）。");
                lines.Add("请右键此工具选择「以管理员身份运行」后重试。");
            }
            if (otherErrors.Count > 0)
            {
                lines.Add($"另有 {otherErrors.Count} 个进程清理失败：");
                lines.AddRange(otherErrors);
            }
            if (lines.Count == 0)
            {
                lines.Add("未发现需要清理的后台 Excel 进程。");
            }

            uint icon = (deniedIds.Count > 0 || otherErrors.Count > 0) ? MB_ICONWARNING : MB_ICONINFORMATION;
            string msg = string.Join(Environment.NewLine, lines) + VisibleNote(visibleCount);
            ShowInfo(msg, icon);

            if (deniedIds.Count > 0) return ExitAccessDenied;
            if (visibleCount > 0) return ExitVisibleRemains;
            return ExitClean;
        }

        /// <summary>
        /// 枚举所有 EXCEL 进程并按「是否拥有可见的顶层窗口」分类。
        ///
        /// 判据说明：不用 Process.MainWindowHandle（该属性在进程刚被枚举、或有其它进程
        /// 刚被结束时可能瞬时返回 0，导致误判），改为用 EnumWindows 枚举所有顶层窗口，
        /// 找出「可见 + 有标题」的窗口归属的进程 PID —— 这些是用户正在使用的 Excel，
        /// 更稳定可靠。没有这类窗口的 EXCEL.EXE 即为后台幽灵进程。
        ///
        /// 无论任务管理器显示为「Microsoft Excel」还是「Excel」，其映像名都是 EXCEL.EXE，
        /// GetProcessesByName("EXCEL") 均能匹配。
        /// </summary>
        private static void ClassifyExcelProcesses(List<Process> ghosts, List<Process> visible)
        {
            Process[] procs;
            try { procs = Process.GetProcessesByName("EXCEL"); }
            catch { return; }

            // 收集「拥有可见顶层窗口」的进程 PID 集合
            var pidsWithVisibleWindow = new HashSet<uint>();
            try
            {
                EnumWindows((hWnd, lParam) =>
                {
                    try
                    {
                        if (!IsWindowVisible(hWnd)) return true;
                        // 排除无标题的隐藏/消息窗口（如 Excel 的 HardwareMonitorWindow）
                        if (GetWindowTextLengthW(hWnd) <= 0) return true;
                        GetWindowThreadProcessId(hWnd, out uint pid);
                        if (pid != 0) pidsWithVisibleWindow.Add(pid);
                    }
                    catch { }
                    return true;
                }, IntPtr.Zero);
            }
            catch { }

            foreach (var p in procs)
            {
                uint pid;
                try { pid = (uint)p.Id; }
                catch { continue; }

                if (pidsWithVisibleWindow.Contains(pid))
                    visible.Add(p);
                else
                    ghosts.Add(p);
            }
        }

        /// <summary>
        /// 结束指定进程并复查是否真的消失。
        /// 返回 true 表示已结束；false 时通过 reason 说明原因（denied=拒绝访问）。
        /// </summary>
        private static bool TryKill(Process p, out string? reason)
        {
            reason = null;
            try
            {
                p.Kill();
                p.WaitForExit(3000);
            }
            catch (Exception ex)
            {
                // 判断是否为「拒绝访问」（权限不足，通常需要提权）
                if (ex is System.ComponentModel.Win32Exception w32 && (uint)w32.NativeErrorCode == 5)
                    reason = "denied";
                else if (ex.Message.IndexOf("denied", StringComparison.OrdinalIgnoreCase) >= 0)
                    reason = "denied";
                else
                    reason = ex.Message;
            }

            // 复查：进程是否真的结束
            Thread.Sleep(150);
            try
            {
                var check = Process.GetProcessById(p.Id);
                if (check != null && !check.HasExited)
                {
                    if (reason == null) reason = "still running";
                    return false;
                }
            }
            catch (ArgumentException)
            {
                // GetProcessById 抛异常 = 进程已不存在 = 已结束
                return true;
            }
            catch { /* 其它异常：以下方判定为准 */ }

            return reason == null;
        }

        private static string VisibleNote(int visibleCount)
        {
            if (visibleCount <= 0) return string.Empty;
            return Environment.NewLine + Environment.NewLine +
                   $"注意：检测到 {visibleCount} 个正在使用的 Excel 窗口，已保留、未关闭。";
        }

        private static void ShowInfo(string text, uint icon = MB_ICONINFORMATION)
        {
            MessageBoxW(IntPtr.Zero, text, Title, MB_OK | icon);
        }
    }
}
