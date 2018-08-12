using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using EnvDTE;
using Thread = System.Threading.Thread;

namespace ConsoleAutoCad.Debugger.VsDebug
{
    internal static class DteHelper
    {
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        /// <summary>
        /// Gets the DTE object from any devenv process.
        /// </summary>
        /// <remarks>
        /// After starting devenv.exe, the DTE object is not ready. We need to try repeatedly and fail after the
        /// timeout.
        /// </remarks>
        /// <param name="processId"></param>
        /// <param name="timeout">Timeout in seconds.</param>
        /// <returns>
        /// Retrieved DTE object or <see langword="null"> if not found.
        /// </see></returns>
        internal static DTE GetDte(int processId, int timeout)
        {
            DTE res = null;
            var startTime = DateTime.Now;

            while (res == null && DateTime.Now.Subtract(startTime).Seconds < timeout)
            {
                Thread.Sleep(1000);
                res = GetDte(processId);
            }

            return res;
        }

        /// <summary>
        /// Gets the DTE object from any devenv process.
        /// </summary>
        /// <param name="processId"></param>
        /// <returns>
        /// Retrieved DTE object or <see langword="null"> if not found.
        /// </see></returns>
        private static DTE GetDte(int processId)
        {
            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;
            DTE runningObject = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                var moniker = new IMoniker[1];
                var numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    runningObject = GetDte(processId, moniker[0], bindCtx, rot);
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return runningObject;
        }

        private static DTE GetDte(int processId, IMoniker moniker, IBindCtx bindCtx, IRunningObjectTable rot)
        {
            string name = null;

            try
            {
                moniker?.GetDisplayName(bindCtx, null, out name);
            }
            catch (UnauthorizedAccessException)
            {
                // Do nothing, there is something in the ROT that we do not have access to.
            }

            var monikerRegex = new Regex(@"!VisualStudio.DTE\.\d+\.\d+\:" + processId, RegexOptions.IgnoreCase);
            if (!String.IsNullOrEmpty(name) && monikerRegex.IsMatch(name))
            {
                object runningObject;
                Marshal.ThrowExceptionForHR(rot.GetObject(moniker, out runningObject));
                return runningObject as DTE;
            }

            return null;
        }
    }
}