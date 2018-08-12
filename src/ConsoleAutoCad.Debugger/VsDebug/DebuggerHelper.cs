using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using EnvDTE;
using NLog;
using Process = EnvDTE.Process;

namespace ConsoleAutoCad.Debugger.VsDebug
{
    /// <summary>
    /// Работа с отладчиком Visual Studio.
    /// </summary>
    public static class DebuggerHelper
    {
        /// <summary>
        /// Присоединение отладчика к процессу с предварительной задержкой.
        /// </summary>
        /// <param name="process">Процесс, к которому будет подключен отладчик.</param>
        /// <param name="milliseconds">Задержка в мс.</param>
        /// <param name="logger">Логгер.</param>
        /// <param name="vsversion">Версия Visual Studio, в которой запускается отладчик.</param>
        [Conditional("DEBUG")]
        public static void AttachDebugger(this System.Diagnostics.Process process, int milliseconds, ILogger logger, string vsversion = VsVersion.Version2017)
        {
            SleepWhenDebug(milliseconds);
            AttachDebugger(process, logger, vsversion);
        }

        /// <summary>
        /// Присоединение отладчика к процессу с предварительной задержкой.
        /// </summary>
        /// <param name="process">Процесс, к которому будет подключен отладчик.</param>
        /// <param name="logger">Логгер.</param>
        /// <param name="vsversion">Версия Visual Studio, в которой запускается отладчик.</param>
        [Conditional("DEBUG")]
        public static void AttachDebugger(this System.Diagnostics.Process process, ILogger logger, string vsversion = VsVersion.Version2017)
        {
            if (!System.Diagnostics.Debugger.IsAttached)
            {
                logger.Debug("Debugger doesn't attached");
                return;
            }
            
            // Reference Visual Studio core
            DTE dte;
            try
            {
                // Работает только с 1 экземпляром запущенной VS 
                //dte = (DTE)Marshal.GetActiveObject(vsversion);

                var vsProcess = System.Diagnostics.Process.GetProcesses().First(x => x.ProcessName.Contains("devenv"));
                dte = GetDte(vsProcess.Id, 10);

                if (dte != null)
                {
                    logger.Debug($"Found Visual Studio {dte.Edition} version {dte.Version}");
                }
            }
            catch (COMException ex)
            {
                logger.Debug(@"Visual studio not found.", ex);
                return;
            }

            // Try loop - Visual Studio may not respond the first time.
            int tryCount = 5;
            while (tryCount-- > 0)
            {
                try
                {
                    var processes = dte.Debugger.LocalProcesses;
                    foreach (var debuggerProcess in processes.Cast<Process>().Where(
                        proc => proc.Name.IndexOf(process.ProcessName, StringComparison.Ordinal) != -1))
                    {
                        debuggerProcess.Attach();
                        logger.Debug($"Attached to process {process.ProcessName} successfully.");
                        break;
                    }

                    System.Threading.Thread.Sleep(50);
                    break;
                }
                catch (COMException e)
                {
                    logger.Debug($"Exception {e.Message}{Environment.NewLine}");
                    System.Threading.Thread.Sleep(25);
                }
            }
        }

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
        private static DTE GetDte(int processId, int timeout)
        {
            DTE res = null;
            var startTime = DateTime.Now;

            while (res == null && DateTime.Now.Subtract(startTime).Seconds < timeout)
            {
                System.Threading.Thread.Sleep(1000);
                res = GetDte(processId);
            }

            return res;
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        /// <summary>
        /// Gets the DTE object from any devenv process.
        /// </summary>
        /// <param name="processId"></param>
        /// <returns>
        /// Retrieved DTE object or <see langword="null"> if not found.
        /// </see></returns>
        private static DTE GetDte(int processId)
        {
            object runningObject = null;

            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];

                    string name = null;

                    try
                    {
                        if (runningObjectMoniker != null)
                        {
                            runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to.
                    }

                    Regex monikerRegex = new Regex(@"!VisualStudio.DTE\.\d+\.\d+\:" + processId, RegexOptions.IgnoreCase);
                    if (!string.IsNullOrEmpty(name) && monikerRegex.IsMatch(name))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                        break;
                    }
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

            return runningObject as DTE;
        }

        /// <summary>
        /// Задержка текущего процесса, если к нему подключен отладчик.
        /// </summary>
        /// <param name="milliseconds">Задержка в мс.</param>
        [Conditional("DEBUG")]
        public static void SleepWhenDebug(int milliseconds)
        {
            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.Out.WriteLine($"Sleep thread for {milliseconds} ms");
                System.Threading.Thread.Sleep(milliseconds);
            }
        }
    }
}