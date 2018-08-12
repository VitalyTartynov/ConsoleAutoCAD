using System.IO;
using System.Reflection;
using ConsoleAutoCad.Runner;
using ConsoleAutoCad.TestPlugin.Output;
using Infrastructure.EmbeddedResources;
using Infrastructure.Universal.Paths;
using NLog;
using NLog.Targets;
using NUnit.Framework;

namespace ConsoleAutoCad.Tests.Runner
{
    [TestFixture]
    [Ignore("Only manual usage allowed!")]
    public class RunnerTests
    {
        const string TestPluginDllName = "ConsoleAutoCad.TestPlugin.dll";

        private IProcessRunner _runner;

        [SetUp]
        public void SetUp()
        {
            var target = new ConsoleTarget {Layout = "${message}"};
            NLog.Config.SimpleConfigurator.ConfigureForTargetLogging(target, LogLevel.Trace);
            var logger = LogManager.GetLogger("Nunit tests");
            
            if (!File.Exists(AcadPaths.Acad2015))
            {
                Assert.Fail($"Console AutoCAD at path '{AcadPaths.Acad2015}' not found!");
            }

            _runner = new AutocadRunner(AcadPaths.Acad2015, logger);
        }

        [TestCase(@"Runner\Samples\lines-3.dwg", 3)]
        [TestCase(@"Runner\Samples\lines-8.dwg", 8)]
        public void LineCounterTests(string inputResourcePath, int expectedCount)
        {
            var pathToPluginDll = Path.Combine(PathHelper.AssemblyDirectory(Assembly.GetExecutingAssembly()), TestPluginDllName);

            using (var dwgFileStream = ResourceLoader.GetEmbeddedResource(Assembly.GetExecutingAssembly(), inputResourcePath))
            {
                var result = _runner.Process<LinesData>(dwgFileStream, pathToPluginDll, command: "LinesCount");

                Assert.That(result, Is.Not.Null);
                Assert.That(result.Count, Is.EqualTo(expected: expectedCount));
            }
        }
    }
}