using System.Threading;
using Autodesk.AutoCAD.Runtime;

namespace ConsoleAutoCad.Debugger
{
    public class DebuggerApplication : IExtensionApplication
    {
        public void Initialize()
        {
            // Можно подключиться к отладчику, вызывав явно диалог выбора
            // System.Diagnostics.Debugger.Launch();

            // эта задержка требуется, чтобы отладчик мог подключиться к процессу и подгрузить нужные pdb файлы
            Thread.Sleep(5000);
        }

        public void Terminate()
        {
        }
    }
}
