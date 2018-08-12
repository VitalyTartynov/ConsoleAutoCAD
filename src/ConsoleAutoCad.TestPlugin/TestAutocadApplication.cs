using System.IO;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using ConsoleAutoCad.Runner;
using ConsoleAutoCad.TestPlugin.Output;
using Infrastructure.Universal;
using Newtonsoft.Json;

namespace ConsoleAutoCad.TestPlugin
{
    public class TestAutocadApplication : IExtensionApplication
    {
        /// <inheritdoc />
        public void Initialize()
        {
            // Nothing to do here.
        }

        /// <inheritdoc />
        public void Terminate()
        {
            // Nothing to do here.
        }

        [CommandMethod("LinesCount")]
        public void LinesCount()
        {
            var result = GetLinesData();
            
            var path = ConsoleAutocad.CurrentFileName + AutocadRunner.OutputFileEnding;
            var serializedResult = JsonConvert.SerializeObject(result);
            File.WriteAllText(path, serializedResult);
        }

        private static LinesData GetLinesData()
        {
            var lines = 0;
            var currentDocument = Application.DocumentManager.MdiActiveDocument;
            var selectionResult = currentDocument.Editor.SelectAll();
            using (var transaction = currentDocument.Database.TransactionManager.StartOpenCloseTransaction())
            {
                var ids = selectionResult?.Value?.GetObjectIds();
                if (ids == null)
                {
                    return null;
                }

                var ts = new TypedSwitch<int>()
                    .Case((Line x) => 1);

                foreach (ObjectId objectId in ids)
                {
                    var currentObject = transaction.GetObject(objectId, OpenMode.ForRead);
                    if (currentObject != null)
                    {
                        lines += ts.Switch(currentObject);
                    }
                }
            }

            return new LinesData {Count = lines};
        }
    }
}