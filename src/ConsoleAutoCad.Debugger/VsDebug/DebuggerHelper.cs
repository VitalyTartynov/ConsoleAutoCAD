using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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
                dte = (DTE)Marshal.GetActiveObject(vsversion);

                //var vsProcess = System.Diagnostics.Process.GetProcesses().First(x => x.ProcessName.Contains("devenv"));
                //dte = DteHelper.GetDte(vsProcess.Id, 10);

                //if (dte != null)
                //{
                //    logger.Debug($"Found Visual Studio {dte.Edition} version {dte.Version}");
                //}
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