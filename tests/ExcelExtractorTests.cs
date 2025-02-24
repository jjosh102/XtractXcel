using System;

using ClosedXML.Excel;

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
        public void ExtractData_ShouldParseExcelIntoCorrectTypes()
        {
            using var stream = CreateTestExcelFile();

            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

            Assert.NotNull(extractedData);
            Assert.Equal(3, extractedData.Count);
        }

        [Fact]
        public void ExtractData_ShouldParseNullableFieldsCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

            Assert.Null(extractedData[1].Age);
            Assert.Null(extractedData[1].Salary);
            Assert.Null(extractedData[0].LastActive);
        }

        [Fact]
        public void ExtractData_ShouldParseNegativeAndZeroValuesCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

            Assert.Equal(0, extractedData[2].Age);
            Assert.Equal(-100.50m, extractedData[2].Salary);
        }

        [Fact]
        public void ExtractData_ShouldParseDatesCorrectly()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

            Assert.Equal(new DateTime(2020, 5, 1), extractedData[0].JoinDate);
            Assert.Equal(new DateTime(2018, 10, 15), extractedData[1].JoinDate);
            Assert.Equal(new DateTime(2022, 1, 1), extractedData[2].JoinDate);
        }


        [Fact]
        public void ExtractData_ShouldHandleMissingColumns()
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

            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

            Assert.NotNull(extractedData);
            Assert.Single(extractedData);
            Assert.Equal(30, extractedData[0].Age);
            Assert.Equal(60000m, extractedData[0].Salary);
            Assert.Null(extractedData[0].Name);
        }


        [Fact]
        public void ExtractData_FromStream_ShouldThrowForNullStream()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelExtractor.ExtractDataFromStream<Person>(null!));
        }

        [Fact]
        public void ExtractDataFromStream_ShouldThrowIfNoPropertiesHaveAttributes()
        {
            using var stream = CreateTestExcelFile();
            var exception = Record.Exception(() => ExcelExtractor.ExtractDataFromStream<NoExcelAttributes>(stream));

            Assert.NotNull(exception);
            Assert.IsType<InvalidOperationException>(exception);
            Assert.Equal($"No properties with {nameof(ExcelColumnAttribute)} found on type NoExcelAttributes", exception.Message);
        }


        [Fact]
        public void ExtractData_FromStream_ShouldReturnValidData()
        {
            using var stream = CreateTestExcelFile();
            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

            Assert.NotNull(extractedData);
            Assert.Equal(3, extractedData.Count);
            Assert.Equal("Alice", extractedData[0].Name);
            Assert.Equal(25, extractedData[0].Age);
        }

        [Fact]
        public void ExtractData_FromFile_ShouldThrowForNullOrWhitespaceFilePath()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelExtractor.ExtractDataFromStream<Person>(null!));
            Assert.Throws<ArgumentException>(() => ExcelExtractor.ExtractDataFromFile<Person>(" "));
        }

        [Fact]
        public void ExtractData_FromFile_ShouldReturnValidData()
        {
            var tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
            try
            {
                using (var stream = CreateTestExcelFile())
                using (var fileStream = File.Create(tempFile))
                {
                    stream.CopyTo(fileStream);
                }

                var extractedData = ExcelExtractor.ExtractDataFromFile<Person>(tempFile);

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
        public void ExtractData_ShouldHandleMultipleFallbackColumns()
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

            var extractedData = ExcelExtractor.ExtractDataFromStream<Person>(stream);

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
