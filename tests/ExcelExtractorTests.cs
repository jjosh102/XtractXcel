using System;
using System.IO;

using ClosedXML.Excel;

using ExcelTransformLoad.Extractor;

using Xunit;

namespace ExcelTransformLoad.Tests
{
    public class ExcelExtractorTests
    {
        private static MemoryStream CreateTestExcelFile()
        {
            var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                // Headers
                worksheet.Cell(1, 1).Value = "Full Name";
                worksheet.Cell(1, 2).Value = "Age";
                worksheet.Cell(1, 3).Value = "Salary";
                worksheet.Cell(1, 4).Value = "Join Date";
                worksheet.Cell(1, 5).Value = "Last Active";

                // Data rows
                worksheet.Cell(2, 1).Value = "Alice";
                worksheet.Cell(2, 2).Value = 25;
                worksheet.Cell(2, 3).Value = 50000.75;
                worksheet.Cell(2, 4).Value = new DateTime(2020, 5, 1);
                worksheet.Cell(2, 5).Clear();

                worksheet.Cell(3, 1).Value = "Bob";
                worksheet.Cell(3, 2).Clear();
                worksheet.Cell(3, 3).Clear();
                worksheet.Cell(3, 4).Value = new DateTime(2018, 10, 15);
                worksheet.Cell(3, 5).Value = new DateTime(2023, 3, 10);

                // Edge cases
                worksheet.Cell(4, 1).Value = "Charlie";
                worksheet.Cell(4, 2).Value = 0;
                worksheet.Cell(4, 3).Value = -100.50;
                worksheet.Cell(4, 4).Value = new DateTime(2022, 1, 1);
                worksheet.Cell(4, 5).Clear();

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            return stream;
        }

        [Fact]
        public void ExtractExtractor_ShouldParseExcelIntoCorrectTypes()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.NotNull(extractedData);
            Assert.Equal(3, extractedData.Count);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseNullableFieldsCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.Null(extractedData[1].Age);
            Assert.Null(extractedData[1].Salary);
            Assert.Null(extractedData[0].LastActive);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseNegativeAndZeroValuesCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.Equal(0, extractedData[2].Age);
            Assert.Equal(-100.50m, extractedData[2].Salary);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseDatesCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.Equal(new DateTime(2020, 5, 1), extractedData[0].JoinDate);
            Assert.Equal(new DateTime(2018, 10, 15), extractedData[1].JoinDate);
            Assert.Equal(new DateTime(2022, 1, 1), extractedData[2].JoinDate);
        }

        [Fact]
        public void ExtractExtractor_ShouldHandleMissingColumns()
        {
            using var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                worksheet.Cell(1, 1).Value = "Age";
                worksheet.Cell(1, 2).Value = "Salary";

                worksheet.Cell(2, 1).Value = 30;
                worksheet.Cell(2, 2).Value = 60000;

                workbook.SaveAs(stream);
            }
            stream.Position = 0;

            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.NotNull(extractedData);
            Assert.Single(extractedData);
            Assert.Equal(30, extractedData[0].Age);
            Assert.Equal(60000m, extractedData[0].Salary);
            Assert.Null(extractedData[0].Name);
        }

        [Fact]
        public void ExtractExtractor_FromStream_ShouldThrowForNullStream()
        {
            var extractor = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1);

            Assert.Throws<ArgumentNullException>(() => extractor.FromStream(null!));
        }

        [Fact]
        public void ExtractExtractor_FromFile_ShouldThrowForNullOrWhitespaceFilePath()
        {
            var extractor = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1);

            Assert.Throws<ArgumentException>(() => extractor.FromFile(string.Empty));
        }

        [Fact]
        public void ExtractDataFromStream_ShouldThrowIfNoPropertiesHaveAttributes()
        {
            using var stream = CreateTestExcelFile();
            var extractor = new ExcelExtractor<NoExcelAttributes>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream);

            var exception = Record.Exception(() => extractor.Extract());

            Assert.NotNull(exception);
            Assert.IsType<InvalidOperationException>(exception);
            Assert.Equal($"No properties with {nameof(ExcelColumnAttribute)} found on type NoExcelAttributes", exception.Message);
        }

        [Fact]
        public void ExtractExtractor_FromStream_ShouldReturnValidData()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.NotNull(extractedData);
            Assert.Equal(3, extractedData.Count);
            Assert.Equal("Alice", extractedData[0].Name);
            Assert.Equal(25, extractedData[0].Age);
        }

        [Fact]
        public void ExtractExtractor_FromFile_ShouldReturnValidData()
        {
            var tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
            try
            {
                using (var stream = CreateTestExcelFile())
                using (var fileStream = File.Create(tempFile))
                {
                    stream.CopyTo(fileStream);
                }

                var extractedData = new ExcelExtractor<Person>()
                    .WithHeader(true)
                    .WithSheetIndex(1)
                    .FromFile(tempFile)
                    .Extract();

                Assert.NotNull(extractedData);
                Assert.Equal(3, extractedData.Count);
                Assert.Equal("Alice", extractedData[0].Name);
                Assert.Equal(25, extractedData[0].Age);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void ExtractExtractor_ShouldHandleMultipleFallbackColumns()
        {
            using var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Employee Age";
                worksheet.Cell(1, 3).Value = "Salary";
                worksheet.Cell(1, 4).Value = "Join Date";
                worksheet.Cell(1, 5).Value = "Last Active";

                worksheet.Cell(2, 1).Value = "Alice";
                worksheet.Cell(2, 2).Value = 25;
                worksheet.Cell(2, 3).Value = 50000.75;
                worksheet.Cell(2, 4).Value = new DateTime(2020, 5, 1);
                worksheet.Cell(2, 5).Clear();

                workbook.SaveAs(stream);
            }
            stream.Position = 0;

            var extractedData = new ExcelExtractor<Person>()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.NotNull(extractedData);
            Assert.Single(extractedData);
            Assert.Equal("Alice", extractedData[0].Name);
        }

        [Fact]
        public void ExtractExtractor_ThrowsWhenModifyingOptionsAfterSourceSet()
        {
            var extractor = new ExcelExtractor<Person>()
                .WithHeader(true)
                .FromStream(new MemoryStream());

            var exception = Record.Exception(() => extractor.WithSheetIndex(1));
            Assert.NotNull(exception);
            Assert.IsType<InvalidOperationException>(exception);
        }

        [Fact]
        public void ExtractExtractor_ThrowsWhenSourceIsSetTwice()
        {
            var extractor = new ExcelExtractor<Person>()
                .WithHeader(true)
                .FromStream(new MemoryStream());

            var exception = Record.Exception(() => extractor.FromFile("path.xlsx"));
            Assert.NotNull(exception);
            Assert.IsType<InvalidOperationException>(exception);
        }

        [Fact]
        public void ExtractExtractor_IncompatibleType_ShouldThrowArgumentException()
        {
            using var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                worksheet.Cell(1, 1).Value = "Grace";
                worksheet.Cell(1, 2).Value = 29;
                worksheet.Cell(1, 3).Value = 29;
                worksheet.Cell(1, 4).Value = 29;
                worksheet.Cell(1, 5).Value = 29;
                worksheet.Cell(1, 6).Value = 29;
                worksheet.Cell(1, 7).Value = 29;

                workbook.SaveAs(stream);
            }

            stream.Position = 0;

            var extractor = new ExcelExtractor<PersonNoHeader>()
                .WithHeader(false)
                .WithSheetIndex(1)
                .FromStream(stream);

            var exception = Record.Exception(() => extractor.Extract());

            Assert.NotNull(exception);
            Assert.IsType<ArgumentException>(exception);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseExcelWithoutHeaders()
        {
            using var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");

                worksheet.Cell(1, 1).Value = "Dave";
                worksheet.Cell(1, 2).Value = 42;
                worksheet.Cell(1, 3).Value = 75000.50;
                worksheet.Cell(1, 4).Value = new DateTime(2019, 3, 15);
                worksheet.Cell(1, 5).Value = new DateTime(2024, 1, 10);

                worksheet.Cell(2, 1).Value = "Eve";
                worksheet.Cell(2, 2).Value = 38;
                worksheet.Cell(2, 3).Value = 82000.25;
                worksheet.Cell(2, 4).Value = new DateTime(2020, 7, 22);
                worksheet.Cell(2, 5).Value = new DateTime(2024, 2, 5);

                workbook.SaveAs(stream);
            }

            stream.Position = 0;

            var extractedData = new ExcelExtractor<PersonNoHeader>()
                .WithHeader(false)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.NotNull(extractedData);
            Assert.Equal(2, extractedData.Count);

            Assert.Equal("Dave", extractedData[0].Name);
            Assert.Equal(42, extractedData[0].Age);
            Assert.Equal(75000.50m, extractedData[0].Salary);
            Assert.Equal(new DateTime(2019, 3, 15), extractedData[0].JoinDate);
            Assert.Equal(new DateTime(2024, 1, 10), extractedData[0].LastActive);

            Assert.Equal("Eve", extractedData[1].Name);
            Assert.Equal(38, extractedData[1].Age);
            Assert.Equal(82000.25m, extractedData[1].Salary);
            Assert.Equal(new DateTime(2020, 7, 22), extractedData[1].JoinDate);
            Assert.Equal(new DateTime(2024, 2, 5), extractedData[1].LastActive);
        }

        [Fact]
        public void ExtractExtractor_WithoutHeader_ShouldHandleMissingValues()
        {
            using var stream = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                worksheet.Cell(1, 1).Value = "Frank";
                worksheet.Cell(1, 2).Clear(); 
                worksheet.Cell(1, 3).Value = 65000.75;
                worksheet.Cell(1, 4).Value = new DateTime(2021, 5, 10);
                worksheet.Cell(1, 5).Clear(); 

                workbook.SaveAs(stream);
            }

            stream.Position = 0;

            var extractedData = new ExcelExtractor<PersonNoHeader>()
                .WithHeader(false)
                .WithSheetIndex(1)
                .FromStream(stream)
                .Extract();

            Assert.NotNull(extractedData);
            Assert.Single(extractedData);
            Assert.Equal("Frank", extractedData[0].Name);
            Assert.Null(extractedData[0].Age);
            Assert.Equal(65000.75m, extractedData[0].Salary);
            Assert.Equal(new DateTime(2021, 5, 10), extractedData[0].JoinDate);
            Assert.Null(extractedData[0].LastActive);
        }

        
    }



    public class Person
    {
        [ExcelColumn("Full Name", "Name", "Employee Name")]
        public string? Name { get; init; }

        [ExcelColumn("Age", "Employee Age")]
        public int? Age { get; init; }

        [ExcelColumn("Salary")]
        public decimal? Salary { get; init; }

        [ExcelColumn("Join Date")]
        public DateTime JoinDate { get; init; }

        [ExcelColumn("Last Active", "Last Activity")]
        public DateTime? LastActive { get; init; }
    }

    public class NoExcelAttributes
    {
        public string? Name { get; init; }
        public int? Age { get; init; }
        public decimal? Salary { get; init; }
        public DateTime JoinDate { get; init; }
        public DateTime? LastActive { get; init; }
    }

    public class PersonNoHeader
    {
        public string? Name { get; init; }
        public int? Age { get; init; }
        public decimal? Salary { get; init; }
        public DateTime JoinDate { get; init; }
        public DateTime? LastActive { get; init; }
    }
}