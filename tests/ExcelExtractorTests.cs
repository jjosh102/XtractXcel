using System;

using ClosedXML.Excel;

using ExcelTransformLoad.Extractor;

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
            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

            Assert.NotNull(extractedData);
            Assert.Equal(3, extractedData.Count);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseNullableFieldsCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

            Assert.Null(extractedData[1].Age);
            Assert.Null(extractedData[1].Salary);
            Assert.Null(extractedData[0].LastActive);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseNegativeAndZeroValuesCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

            Assert.Equal(0, extractedData[2].Age);
            Assert.Equal(-100.50m, extractedData[2].Salary);
        }

        [Fact]
        public void ExtractExtractor_ShouldParseDatesCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

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

            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

            Assert.NotNull(extractedData);
            Assert.Single(extractedData);
            Assert.Equal(30, extractedData[0].Age);
            Assert.Equal(60000m, extractedData[0].Salary);
            Assert.Null(extractedData[0].Name);
        }


        [Fact]
        public void ExtractExtractor_FromStream_ShouldThrowForNullStream()
        {
            Assert.Throws<ArgumentNullException>(() => new StreamExcelExtractor<Person>(null!));
        }

        [Fact]
        public void ExtractExtractor_FromFile_ShouldThrowForNullOrWhitespaceFilePath()
        {
            Assert.Throws<ArgumentException>(() => new FileExcelExtractor<Person>(string.Empty));
        }

        [Fact]
        public void ExtractDataFromStream_ShouldThrowIfNoPropertiesHaveAttributes()
        {
            using var stream = CreateTestExcelFile();
            var extractor = new StreamExcelExtractor<NoExcelAttributes>(stream);
            var exception = Record.Exception(() => extractor.Extract());

            Assert.NotNull(exception);
            Assert.IsType<InvalidOperationException>(exception);
            Assert.Equal($"No properties with {nameof(ExcelColumnAttribute)} found on type NoExcelAttributes", exception.Message);
        }


        [Fact]
        public void ExtractExtractor_FromStream_ShouldReturnValidData()
        {
            using var stream = CreateTestExcelFile();
            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

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

                var extractor = new FileExcelExtractor<Person>(tempFile);
                var extractedData = extractor.Extract();

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

                // Headers
                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Employee Age";
                worksheet.Cell(1, 3).Value = "Salary";
                worksheet.Cell(1, 4).Value = "Join Date";
                worksheet.Cell(1, 5).Value = "Last Active";

                // Data rows
                worksheet.Cell(2, 1).Value = "Alice";
                worksheet.Cell(2, 2).Value = 25;
                worksheet.Cell(2, 3).Value = 50000.75;
                worksheet.Cell(2, 4).Value = new DateTime(2020, 5, 1);
                worksheet.Cell(2, 5).Clear();

                workbook.SaveAs(stream);
            }
            stream.Position = 0;

            var extractor = new StreamExcelExtractor<Person>(stream);
            var extractedData = extractor.Extract();

            Assert.NotNull(extractedData);
            Assert.Single(extractedData);
            Assert.Equal("Alice", extractedData[0].Name);
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
}
