using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;
using XtractXcel;

namespace ExcelTransformLoad.Benchmarks
{
    // dotnet commands
    // dotnet run --framework net8.0 net9.0 --configuration Release --no-debug
    // dotnet run --configuration Release --no-debug
    public class Program
    {
        public static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<ExcelExtractionBenchmark>();
            Console.WriteLine(summary);
        }
    }
}