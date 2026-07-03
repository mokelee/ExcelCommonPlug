using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelCommonTools.Core
{
    /// <summary>
    /// COM 对象释放辅助工具。
    /// 提供安全的 COM RCW 释放方法，避免 Excel COM 互操作中的内存泄漏。
    /// </summary>
    internal static class ComHelper
    {
        /// <summary>
        /// 安全释放单个 COM 对象。
        /// 如果对象为 null 或不是 COM 对象，则不做任何操作。
        /// </summary>
        /// <param name="obj">要释放的 COM 对象</param>
        public static void Release(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// 安全释放单个 COM 对象（带上下文日志）。
        /// 释放失败时记录调试日志，不抛出异常。
        /// </summary>
        /// <param name="obj">要释放的 COM 对象</param>
        /// <param name="context">释放失败的上下文描述</param>
        public static void Release(object obj, string context)
        {
            try
            {
                Release(obj);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ComHelper] 释放 COM 对象失败 ({context}): {ex.Message}");
            }
        }
    }
}
