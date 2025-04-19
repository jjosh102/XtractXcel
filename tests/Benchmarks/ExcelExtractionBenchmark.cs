using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;

using ClosedXML.Excel;

using XtractXcel;

namespace ExcelTransformLoad.Benchmarks
{
    [MemoryDiagnoser]
    public class ExcelExtractionBenchmark
    {
        private MemoryStream? _largeExcelStream;
        private MemoryStream? _mediumExcelStream;
        private MemoryStream? _smallExcelStream;

        [GlobalSetup]
        public void Setup()
        {
            _smallExcelStream = CreateTestExcelFile(100);
            _mediumExcelStream = CreateTestExcelFile(1000);
            _largeExcelStream = CreateTestExcelFile(10000);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            _smallExcelStream?.Dispose();
            _mediumExcelStream?.Dispose();
            _largeExcelStream?.Dispose();
        }

        [Benchmark]
        public List<Person> SmallFile_AttributeMapping()
        {
            EnsureStreamIsNotNull(_smallExcelStream);

            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(_smallExcelStream!)
                .Extract<Person>();
        }

        [Benchmark]
        public List<Person> SmallFile_ManualMapping()
        {
            EnsureStreamIsNotNull(_smallExcelStream);

            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(_smallExcelStream!)
                .ExtractWithManualMapping(row => new Person
                {
                    Name = row.Cell(1).IsEmpty() ? null : row.Cell(1).GetString(),
                    Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                    Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                    JoinDate = row.Cell(4).GetDateTime(),
                    LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
                });
        }

        [Benchmark]
        public List<NoExcelAttributes> SmallFile_ManualMapping_NoAttributes()
        {
            EnsureStreamIsNotNull(_smallExcelStream);

            return new ExcelExtractor()
            .WithHeader(true)
            .WithWorksheetIndex(1)
            .FromStream(_smallExcelStream!)
            .ExtractWithManualMapping(row => new NoExcelAttributes
            {
                Name = row.Cell(1).IsEmpty() ? null : row.Cell(1).GetString(),
                Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                JoinDate = row.Cell(4).GetDateTime(),
                LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
            });
        }

        [Benchmark]
        public List<Person> MediumFile_AttributeMapping()
        {
            EnsureStreamIsNotNull(_mediumExcelStream);
            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(_mediumExcelStream!)
                .Extract<Person>();
        }

        [Benchmark]
        public List<Person> MediumFile_ManualMapping()
        {
            EnsureStreamIsNotNull(_mediumExcelStream);
            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(_mediumExcelStream!)
                .ExtractWithManualMapping(row => new Person
                {
                    Name = row.Cell(1).IsEmpty() ? null : row.Cell(1).GetString(),
                    Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                    Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                    JoinDate = row.Cell(4).GetDateTime(),
                    LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
                });
        }

        [Benchmark]
        public List<Person> LargeFile_AttributeMapping()
        {
            EnsureStreamIsNotNull(_largeExcelStream);
            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(_largeExcelStream!)
                .Extract<Person>();
        }

        [Benchmark]
        public List<Person> LargeFile_ManualMapping()
        {
            EnsureStreamIsNotNull(_largeExcelStream);
            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(_largeExcelStream!)
                .ExtractWithManualMapping(row => new Person
                {
                    Name = row.Cell(1).IsEmpty() ? null : row.Cell(1).GetString(),
                    Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                    Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                    JoinDate = row.Cell(4).GetDateTime(),
                    LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
                });
        }

        [Benchmark]
        public List<Person> ManyColumns_AttributeMapping()
        {
            var stream = CreateTestExcelFileWithManyColumns();
            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(stream)
                .Extract<Person>();
        }

        [Benchmark]
        public List<Person> ManyColumns_ManualMapping()
        {
            var stream = CreateTestExcelFileWithManyColumns();
            return new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(stream)
                .ExtractWithManualMapping(row => new Person
                {
                    Name = row.Cell(1).IsEmpty() ? null : row.Cell(1).GetString(),
                    Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                    Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                    JoinDate = row.Cell(4).GetDateTime(),
                    LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
                });
        }

        private static MemoryStream CreateTestExcelFile(int rowCount)
        {
            var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                worksheet.Cell(1, 1).Value = "Full Name";
                worksheet.Cell(1, 2).Value = "Age";
                worksheet.Cell(1, 3).Value = "Salary";
                worksheet.Cell(1, 4).Value = "Join Date";
                worksheet.Cell(1, 5).Value = "Last Active";

                var random = new Random(42);

                for (int i = 0; i < rowCount; i++)
                {
                    int row = i + 2;

                    worksheet.Cell(row, 1).Value = $"Person {i}";

                    if (random.NextDouble() > 0.1)
                    {
                        worksheet.Cell(row, 2).Value = 20 + random.Next(40);
                    }
                    else
                    {
                        worksheet.Cell(row, 2).Clear();
                    }

                    if (random.NextDouble() > 0.2)
                    {
                        worksheet.Cell(row, 3).Value = 30000 + random.Next(70000) + random.NextDouble();
                    }
                    else
                    {
                        worksheet.Cell(row, 3).Clear();
                    }

                    var joinYear = 2015 + random.Next(9);
                    var joinMonth = 1 + random.Next(12);
                    var joinDay = 1 + random.Next(28);
                    worksheet.Cell(row, 4).Value = new DateTime(joinYear, joinMonth, joinDay);

                    if (random.NextDouble() > 0.3)
                    {
                        var lastYear = joinYear + random.Next(3);
                        var lastMonth = 1 + random.Next(12);
                        var lastDay = 1 + random.Next(28);
                        worksheet.Cell(row, 5).Value = new DateTime(lastYear, lastMonth, lastDay);
                    }
                    else
                    {
                        worksheet.Cell(row, 5).Clear();
                    }
                }

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            return stream;
        }

        private static MemoryStream CreateTestExcelFileWithManyColumns()
        {
            var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                int totalColumns = 50;

                worksheet.Cell(1, 1).Value = "Full Name";
                worksheet.Cell(1, 2).Value = "Age";
                worksheet.Cell(1, 3).Value = "Salary";
                worksheet.Cell(1, 4).Value = "Join Date";
                worksheet.Cell(1, 5).Value = "Last Active";

                for (int i = 6; i <= totalColumns; i++)
                {
                    worksheet.Cell(1, i).Value = $"Extra Column {i}";
                }

                var random = new Random(42);

                for (int i = 0; i < 100; i++)
                {
                    int row = i + 2;

                    worksheet.Cell(row, 1).Value = $"Person {i}";
                    worksheet.Cell(row, 2).Value = 20 + random.Next(40);
                    worksheet.Cell(row, 3).Value = 30000 + random.Next(70000) + random.NextDouble();
                    worksheet.Cell(row, 4).Value = new DateTime(2020, 1 + random.Next(12), 1 + random.Next(28));

                    if (random.NextDouble() > 0.3)
                    {
                        worksheet.Cell(row, 5).Value = new DateTime(2022, 1 + random.Next(12), 1 + random.Next(28));
                    }
                    else
                    {
                        worksheet.Cell(row, 5).Clear();
                    }

                    for (int col = 6; col <= totalColumns; col++)
                    {
                        if (random.NextDouble() > 0.5)
                        {
                            worksheet.Cell(row, col).Value = $"Data {i}-{col}";
                        }
                        else
                        {
                            worksheet.Cell(row, col).Value = random.NextDouble() * 1000;
                        }
                    }
                }

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            return stream;
        }

        private void EnsureStreamIsNotNull(MemoryStream? stream)
        {
            if (stream is not null)
            {
                stream.Position = 0;
            }
        }
    }
}