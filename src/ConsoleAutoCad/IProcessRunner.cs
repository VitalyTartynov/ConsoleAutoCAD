using System.IO;

namespace ConsoleAutoCad
{
    public interface IProcessRunner
    {
        /// <summary>
        /// Консольный обработчик DWG файла.
        /// </summary>
        /// <param name="binaryStream">Поток данных обрабатываемого DWG файла.</param>
        /// <param name="pathToPluginDll">Абсолютный путь к dll плагина, который содержит команды для обработки DWG файла.</param>
        /// <param name="command">Выполняемая команда, которую добавляет dll плагина.</param>
        /// <param name="showConsoleWindow">Показывать ли окно консольного AutoCAD пользователю.</param>
        /// <returns>Типизированный результат обработки.</returns>
        TOutputContent Process<TOutputContent>(Stream binaryStream, string pathToPluginDll, string command, bool showConsoleWindow = false)
            where TOutputContent : class, IOutputContent;

        /// <summary>
        /// Консольный обработчик DWG файла.
        /// </summary>
        /// <param name="pathToDwgFile">Абсолютный путь к обрабатываемому DWG файлу.</param>
        /// <param name="pathToPluginDll">Абсолютный путь к dll плагина, который содержит команды для обработки DWG файла.</param>
        /// <param name="command">Выполняемая команда, которую добавляет dll плагина.</param>
        /// <param name="showConsoleWindow">Показывать ли окно консольного AutoCAD пользователю.</param>
        /// <returns>Типизированный результат обработки.</returns>
        TOutputContent Process<TOutputContent>(string pathToDwgFile, string pathToPluginDll, string command, bool showConsoleWindow = false)
            where TOutputContent : class, IOutputContent;
    }
}