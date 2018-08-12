using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using ConsoleAutoCad.Debugger.VsDebug;
using Infrastructure.Universal.Paths;
using Infrastructure.Universal.Temps;
using Newtonsoft.Json;
using NLog;

namespace ConsoleAutoCad.Runner
{
    /// <summary>
    /// Ответственный за запуск процесса AutoCAD и получение результата обработки файла.
    /// </summary>
    public class AutocadRunner : IProcessRunner
    {
        /// <summary>
        /// Окончание названия файла с результатами обработки. Формируется по принципу: "{FileName}{OutputFileEnding}".
        /// </summary>
        public const string OutputFileEnding = "output.json";

        /// <summary>
        /// Длительность ожидания завершения работы команды.
        /// </summary>
        private const int DefaultTimeout = 30000;

        private readonly Encoding _commandScriptEncoding = Encoding.GetEncoding(1251);
        
        private readonly string _autocadPath;
        private readonly ILogger _log;
        private readonly int _millisecondsToExit;
        private readonly bool _attachDebuggerIfNeeded;
        private readonly string _vsversion;

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="pathToAcCoreConsole">Путь к консольному AutoCAD (<c>accoreconsole.exe</c>).</param>
        /// <param name="log">Логгер, в который будет выводиться информация о работе консольного AutoCAD.</param>
        /// <param name="millisecondsToExit">Длительность ожидания завершения работы команды. Если команда не успеет 
        /// выполниться за указанное время, выполнение программы продолжится после указанного количества миллисекунд;
        /// команда же продолжит выполняться в фоновом режиме. Чтобы дождаться окончания выполнения команды 
        /// во что бы то ни стало, используйте значение <see cref="Timeout.Infinite"/>.</param>
        /// <param name="attachDebuggerIfNeeded">Указывает, требуется ли подключать отладчик к консольному AutoCAD,
        /// если текущий процесс находится под отладкой.</param>
        /// <param name="vsVersion">Версия Visual Studio, которая будет использоваться для отладки.</param>
        public AutocadRunner(string pathToAcCoreConsole, ILogger log, int millisecondsToExit,
            bool attachDebuggerIfNeeded = true, string vsVersion = VsVersion.Version2017)
        {
            _autocadPath = pathToAcCoreConsole;
            _log = log;
            _millisecondsToExit = millisecondsToExit;
            _attachDebuggerIfNeeded = attachDebuggerIfNeeded;
            _vsversion = vsVersion;
        }

        /// <summary>
        /// Ответственный за запуск процесса AutoCAD и получение результата обработки файла.
        /// </summary>
        /// <param name="autocadPath">Путь к консольному AutoCAD.</param>
        /// <param name="log">Логгер.</param>
        /// <param name="vsversion">Версия Visual Studio (см. VsVersion class).</param>
        public AutocadRunner(string autocadPath, ILogger log, string vsversion = VsVersion.Version2017)
            : this(autocadPath, log, DefaultTimeout, true, vsversion)
        {
        }

        /// <summary>
        /// Консольный обработчик DWG файла.
        /// </summary>
        /// <param name="binaryStream">Поток данных обрабатываемого DWG файла.</param>
        /// <param name="pathToPluginDll">Абсолютный путь к dll плагина, который содержит команды для обработки DWG файла.</param>
        /// <param name="command">Выполняемая команда, которую добавляет dll плагина.</param>
        /// <param name="showConsoleWindow">Показывать ли окно консольного AutoCAD пользователю.</param>
        /// <returns>Типизированный результат обработки.</returns>
        public TOutputContent Process<TOutputContent>(Stream binaryStream, string pathToPluginDll, string command, bool showConsoleWindow = false)
            where TOutputContent : class, IOutputContent
        {
            using (var tempDwgFile = TempFile.Create(".dwg", binaryStream))
            {
                return Process<TOutputContent>(tempDwgFile.Path, pathToPluginDll, command, showConsoleWindow);
            }
        }

        /// <summary>
        /// Открывает указанный чертёж при помощи консольного AutoCAD, загружает указанный плагин и выполняет указанную команду.
        /// </summary>
        /// <param name="binaryStream">Поток с данными чертежа, который требуется обработать.</param>
        /// <param name="pathToPluginDll">Полный путь к сборке с плагином, при помощи которой будет обрабатываться чертёж.</param>
        /// <param name="command">Команда, которую требуется запустить для обработки чертежа.</param>
        /// <param name="showConsoleWindow">Флаг, указывающий, нужно ли отображать окно консольгого AutoCAD после запуска.</param>
        public void RunCommmand(Stream binaryStream, string pathToPluginDll, string command, bool showConsoleWindow = false)
        {
            using (var tempDwgFile = TempFile.Create(".dwg", binaryStream))
            {
                RunCommand(tempDwgFile.Path, pathToPluginDll, command, showConsoleWindow);
            }
        }

        /// <summary>
        /// Открывает указанный чертёж при помощи консольного AutoCAD, загружает указанный плагин и выполняет указанную команду.
        /// После выполнения команды этот метод читает данные, порождённые выполненной командой, 
        /// </summary>
        /// <param name="pathToDwgFile">Абсолютный путь к обрабатываемому DWG файлу</param>
        /// <param name="pathToPluginDll">Абсолютный путь к dll плагина, который содержит команды для обработки DWG файла</param>
        /// <param name="command">Выполняемая команда, которую добавляет dll плагина</param>
        /// <param name="showConsoleWindow">Показывать ли окно консольного AutoCAD пользователю</param>
        /// <returns>Типизированный результат обработки</returns>
        public TOutputContent Process<TOutputContent>(string pathToDwgFile, string pathToPluginDll, string command, bool showConsoleWindow = false)
            where TOutputContent : class, IOutputContent
        {
            RunCommand(pathToDwgFile, pathToPluginDll, command, showConsoleWindow);

            var resultPath = $"{pathToDwgFile}{OutputFileEnding}";
            if (!File.Exists(resultPath))
            {
                _log.Error($"Result file for parse '{pathToDwgFile}' not found!");
                return null;
            }

            var acadContent = JsonConvert.DeserializeObject<TOutputContent>(File.ReadAllText(resultPath));
            if (File.Exists(resultPath))
            {
                File.Delete(resultPath);
                _log.Debug($"Deleted result temp file '{resultPath}'");
            }

            return acadContent;
        }

        /// <summary>
        /// Открывает указанный чертёж при помощи консольного AutoCAD, загружает указанный плагин и выполняет указанную команду.
        /// </summary>
        /// <param name="pathToDwgFile">Полный путь к обрабатываемому чертежу.</param>
        /// <param name="pathToPluginDll">Полный путь к сборке с плагином, при помощи которой будет обрабатываться чертёж.</param>
        /// <param name="command">Команда, которую требуется запустить для обработки чертежа.</param>
        /// <param name="showConsoleWindow">Флаг, указывающий, нужно ли отображать окно консольгого AutoCAD после запуска.</param>
        public void RunCommand(string pathToDwgFile, string pathToPluginDll, string command,
            bool showConsoleWindow = false)
        {
            var scriptPath = CreateLoadingScript(pathToPluginDll, command);
            var arguments = $" /i \"{pathToDwgFile}\" /s \"{scriptPath}\"";
            var process = new Process
            {
                StartInfo =
                {
                    FileName = _autocadPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = !showConsoleWindow,
                    WorkingDirectory = Path.GetDirectoryName(pathToDwgFile)
                }
            };

            process.Start();

            if (_attachDebuggerIfNeeded)
            {
                process.AttachDebugger(1000, _log, _vsversion);
            }

            process.WaitForExit(_millisecondsToExit);

            Thread.Sleep(10);

            if (File.Exists(scriptPath))
            {
                File.Delete(scriptPath);
                _log.Debug($"Deleted loading script temp file '{scriptPath}'");
            }
        }

        /// <summary>
        /// Создание скрипта для загрузки плагина в AutoCAD и выполнения команды.
        /// </summary>
        /// <param name="pathToPluginDll">Абсолютный путь к dll плагина.</param>
        /// <param name="command">Выполняемая команда, которую добавляет dll плагина.</param>
        /// <returns>Абсолютный путь к созданому файлу скрипта.</returns>
        public string CreateLoadingScript(string pathToPluginDll, string command)
        {
            var filePath = GetTempFile(".scr");

            var contentBuilder = new StringBuilder();

            // DISABLE checking trusted locations for 2015+ AutoCADs
            contentBuilder.AppendLine("SECURELOAD 0");

            if (_attachDebuggerIfNeeded && System.Diagnostics.Debugger.IsAttached)
            {
                var pathToDebuggerDll = Path.Combine(PathHelper.AssemblyDirectory(Assembly.GetExecutingAssembly()), "ConsoleAutoCad.Debugger.dll");
                if (!File.Exists(pathToDebuggerDll))
                {
                    throw new ApplicationException($"Debug plugin was not found! File '{pathToDebuggerDll}' doesn't exist!");
                }

                contentBuilder.AppendLine($"netload \"{pathToDebuggerDll}\"");
            }

            contentBuilder.AppendLine($"netload \"{pathToPluginDll}\"");
            contentBuilder.AppendLine($"{command}");

            var content = contentBuilder.ToString();
            File.WriteAllText(filePath, content, _commandScriptEncoding);
            _log.Debug($"Created loading script temp file '{filePath}'. Content: '{content}'");
            
            return filePath;
        }

        /// <summary>
        /// Путь к временному файлу с заданным расширением.
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        private string GetTempFile(string extension)
        {
            var path = Path.GetTempFileName();
            var ext = Path.GetExtension(path);
            path = path.Replace(ext, extension);

            return path;
        }
    }
}