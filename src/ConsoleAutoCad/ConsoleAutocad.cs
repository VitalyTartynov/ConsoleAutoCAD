using System.IO;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using ConsoleAutoCad.Runner;
using Newtonsoft.Json;

namespace ConsoleAutoCad
{
    /// <summary>
    /// Служебные поля и методы для работы с консольным автокадом
    /// </summary>
    public static class ConsoleAutocad
    {
        /// <summary>
        /// Текущая база данных
        /// </summary>
        public static Database CurrentDatabase => Application.DocumentManager.MdiActiveDocument.Database;

        /// <summary>
        /// Имя текущей базы данных
        /// </summary>
        public static string CurrentFileName => CurrentDatabase.OriginalFileName;

        /// <summary>
        /// Сохранение результатов для передачи в тесты
        /// </summary>
        /// <param name="result">Интерфейсный объект с результатами. Должен быть сериализуемым</param>
        public static void SaveResultData(IOutputContent result)
        {
            var path = CurrentFileName + AutocadRunner.OutputFileEnding;
            var serializedResult = JsonConvert.SerializeObject(result);
            File.WriteAllText(path, serializedResult);
        }
    }
}