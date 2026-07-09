using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;

namespace ExcelCommonTools.Core
{
    /// <summary>
    /// 简易日志系统。后台线程消费队列写入文件，不阻塞调用方。
    /// 在 AddIn.AutoOpen 中调用 Logger.Init() 初始化。
    /// </summary>
    internal static class Logger
    {
        private static readonly ConcurrentQueue<string> _queue = new ConcurrentQueue<string>();
        private static string _logDirectory;
        private static string _currentFilePath;
        private static Thread _writerThread;
        private static volatile bool _initialized;
        private static readonly object _initLock = new object();
        private static readonly ManualResetEventSlim _signal = new ManualResetEventSlim(false);
        private static readonly long MaxFileSize = 5 * 1024 * 1024; // 5MB
        private static readonly int MaxFileCount = 5;
        private static readonly int MaxFileAgeDays = 7;

        public static void Init(string logDirectory = null)
        {
            if (_initialized) return;
            lock (_initLock)
            {
                if (_initialized) return;

                _logDirectory = string.IsNullOrWhiteSpace(logDirectory)
                    ? Path.Combine(Path.GetDirectoryName(typeof(Logger).Assembly.Location ?? ""), "logs")
                    : logDirectory;

                try
                {
                    if (!Directory.Exists(_logDirectory))
                        Directory.CreateDirectory(_logDirectory);

                    CleanupOldFiles();
                    _currentFilePath = CreateNewLogFilePath();

                    _writerThread = new Thread(ConsumerLoop)
                    {
                        IsBackground = true,
                        Name = "ExcelCommonTools.Logger"
                    };
                    _writerThread.Start();
                    _initialized = true;

                    // 写诊断文件确认初始化成功
                    try
                    {
                        var diagPath = Path.Combine(_logDirectory, "logger_init_ok.txt");
                        File.WriteAllText(diagPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Logger.Init OK, logDir={_logDirectory}, filePath={_currentFilePath}");
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    // 初始化失败，尝试写桌面诊断
                    try
                    {
                        var diagPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                            "logger_init_error.txt");
                        File.WriteAllText(diagPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Logger.Init failed: logDir={_logDirectory}\r\n{ex}");
                    }
                    catch { }
                    throw; // 让调用方也能捕获
                }
            }
        }

        public static void Debug(string module, string message)
        {
            Enqueue("DEBUG", module, message, null);
        }

        public static void Info(string module, string message)
        {
            Enqueue("INFO", module, message, null);
        }

        public static void Warn(string module, string message)
        {
            Enqueue("WARN", module, message, null);
        }

        public static void Error(string module, string message, Exception ex = null)
        {
            Enqueue("ERROR", module, message, ex);
        }

        private static void Enqueue(string level, string module, string message, Exception ex)
        {
            if (!_initialized) return;
            if (_queue.Count >= 10000) return; // 防止内存爆

            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] [{module}] {message}";
            if (ex != null)
                line += Environment.NewLine + ex.ToString();

            _queue.Enqueue(line);
            _signal.Set();
        }

        private static void ConsumerLoop()
        {
            while (true)
            {
                try
                {
                    _signal.Wait(TimeSpan.FromSeconds(1));
                    _signal.Reset();

                    while (_queue.TryDequeue(out var line))
                    {
                        WriteLineToFile(line);
                    }
                }
                catch { }
            }
        }

        private static void WriteLineToFile(string line)
        {
            try
            {
                if (_currentFilePath != null && File.Exists(_currentFilePath))
                {
                    var fi = new FileInfo(_currentFilePath);
                    if (fi.Length >= MaxFileSize)
                        _currentFilePath = CreateNewLogFilePath();
                }
                if (_currentFilePath == null)
                    _currentFilePath = CreateNewLogFilePath();

                File.AppendAllText(_currentFilePath, line + Environment.NewLine);
            }
            catch { }
        }

        private static string CreateNewLogFilePath()
        {
            return Path.Combine(_logDirectory, $"ExcelCommonTools_{DateTime.Now:yyyyMMdd_HHmmss}.log");
        }

        private static void CleanupOldFiles()
        {
            try
            {
                var dir = new DirectoryInfo(_logDirectory);
                if (!dir.Exists) return;

                var logFiles = dir.GetFiles("ExcelCommonTools_*.log")
                    .OrderBy(f => f.CreationTime).ToList();

                var cutoff = DateTime.Now.AddDays(-MaxFileAgeDays);
                foreach (var file in logFiles.Where(f => f.CreationTime < cutoff).ToList())
                {
                    try { file.Delete(); logFiles.Remove(file); } catch { }
                }
                while (logFiles.Count > MaxFileCount)
                {
                    try { logFiles[0].Delete(); } catch { }
                    logFiles.RemoveAt(0);
                }
            }
            catch { }
        }
    }
}
