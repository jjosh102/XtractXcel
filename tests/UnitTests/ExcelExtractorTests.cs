using ClosedXML.Excel;
using ExcelTransformLoad.Extractor;
using ExcelTransformLoad.UnitTests.TestHelpers;

namespace ExcelTransformLoad.UnitTests;

public class ExcelExtractorTests
{
    [Fact]
    public void ExcelExtractor_ShouldParseExcelIntoCorrectTypes()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
    }

    [Fact]
    public void ExcelExtractor_ShouldParseNullableFieldsCorrectly()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.Null(extractedData[1].Age);
        Assert.Null(extractedData[1].Salary);
        Assert.Null(extractedData[0].LastActive);
    }

    [Fact]
    public void ExcelExtractor_ShouldParseNegativeAndZeroValuesCorrectly()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.Equal(0, extractedData[2].Age);
        Assert.Equal(-100.50m, extractedData[2].Salary);
    }

    [Fact]
    public void ExcelExtractor_ShouldParseDatesCorrectly()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.Equal(new DateTime(2020, 5, 1), extractedData[0].JoinDate);
        Assert.Equal(new DateTime(2018, 10, 15), extractedData[1].JoinDate);
        Assert.Equal(new DateTime(2022, 1, 1), extractedData[2].JoinDate);
    }

    [Fact]
    public void ExcelExtractor_ShouldHandleMissingColumns()
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

        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.NotNull(extractedData);
        Assert.Single(extractedData);
        Assert.Equal(30, extractedData[0].Age);
        Assert.Equal(60000m, extractedData[0].Salary);
        Assert.Null(extractedData[0].Name);
    }

    [Fact]
    public void ExcelExtractor_FromStream_ShouldThrowForNullStream()
    {
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1);

        Assert.Throws<ArgumentNullException>(() => extractor.FromStream(null!));
    }

    [Fact]
    public void ExcelExtractor_FromFile_ShouldThrowForNullOrWhitespaceFilePath()
    {
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1);

        Assert.Throws<ArgumentException>(() => extractor.FromFile(string.Empty));
    }

    [Fact]
    public void ExtractDataFromStream_ShouldThrowIfNoPropertiesHaveAttributes()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream);

        var exception = Record.Exception(() => extractor.Extract<NoExcelAttributes>());

        Assert.NotNull(exception);
        Assert.IsType<InvalidOperationException>(exception);
        Assert.Equal($"No properties with {nameof(ExcelColumnAttribute)} found on type NoExcelAttributes",
            exception.Message);
    }

    [Fact]
    public void ExcelExtractor_FromStream_ShouldReturnValidData()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
        Assert.Equal("Alice", extractedData[0].Name);
        Assert.Equal(25, extractedData[0].Age);
    }

    [Fact]
    public void ExcelExtractor_FromFile_ShouldReturnValidData()
    {
        var tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
        try
        {
            using (var stream = TestExcelGenerator.CreateTestExcelFile())
            using (var fileStream = File.Create(tempFile))
            {
                stream.CopyTo(fileStream);
            }

            var extractedData = new ExcelExtractor()
                .WithHeader(true)
                .WithSheetIndex(1)
                .FromFile(tempFile)
                .Extract<Person>();

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
    public void ExcelExtractor_ShouldHandleMultipleFallbackColumns()
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

        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<Person>();

        Assert.NotNull(extractedData);
        Assert.Single(extractedData);
        Assert.Equal("Alice", extractedData[0].Name);
    }

    [Fact]
    public void ExcelExtractor_ThrowsWhenModifyingOptionsAfterSourceSet()
    {
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .FromStream(new MemoryStream());

        var exception = Record.Exception(() => extractor.WithSheetIndex(1));
        Assert.NotNull(exception);
        Assert.IsType<InvalidOperationException>(exception);
    }

    [Fact]
    public void ExcelExtractor_ThrowsWhenSourceIsSetTwice()
    {
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .FromStream(new MemoryStream());

        var exception = Record.Exception(() => extractor.FromFile("path.xlsx"));
        Assert.NotNull(exception);
        Assert.IsType<InvalidOperationException>(exception);
    }

    [Fact]
    public void ExcelExtractor_IncompatibleType_ShouldThrowArgumentException()
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

        var extractor = new ExcelExtractor()
            .WithHeader(false)
            .WithSheetIndex(1)
            .FromStream(stream);

        var exception = Record.Exception(() => extractor.Extract<PersonNoHeader>());

        Assert.NotNull(exception);
        Assert.IsType<ArgumentException>(exception);
    }

    [Fact]
    public void ExcelExtractor_ShouldParseExcelWithoutHeaders()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFileWithNoHeader();

        var extractedData = new ExcelExtractor()
            .WithHeader(false)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonNoHeader>();

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
    public void ExcelExtractor_WithoutHeader_ShouldHandleMissingValues()
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

        var extractedData = new ExcelExtractor()
            .WithHeader(false)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonNoHeader>();

        Assert.NotNull(extractedData);
        Assert.Single(extractedData);
        Assert.Equal("Frank", extractedData[0].Name);
        Assert.Null(extractedData[0].Age);
        Assert.Equal(65000.75m, extractedData[0].Salary);
        Assert.Equal(new DateTime(2021, 5, 10), extractedData[0].JoinDate);
        Assert.Null(extractedData[0].LastActive);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldExtractDataCorrectly()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new Person
            {
                Name = row.Cell(1).GetString(),
                Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                JoinDate = row.Cell(4).GetDateTime(),
                LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
            });

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
        Assert.Equal("Alice", extractedData[0].Name);
        Assert.Equal(25, extractedData[0].Age);
        Assert.Equal(50000.75m, extractedData[0].Salary);
        Assert.Equal(new DateTime(2020, 5, 1), extractedData[0].JoinDate);
        Assert.Null(extractedData[0].LastActive);

        Assert.Equal("Bob", extractedData[1].Name);
        Assert.Null(extractedData[1].Age);
        Assert.Null(extractedData[1].Salary);
        Assert.Equal(new DateTime(2018, 10, 15), extractedData[1].JoinDate);
        Assert.Equal(new DateTime(2023, 3, 10), extractedData[1].LastActive);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldWorkWithoutHeader()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFileWithNoHeader();
        var extractedData = new ExcelExtractor()
            .WithHeader(false)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new PersonNoHeader
            {
                Name = row.Cell(1).GetString(),
                Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                JoinDate = row.Cell(4).GetDateTime(),
                LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
            });


        Assert.NotNull(extractedData);
        Assert.Equal(2, extractedData.Count);
        Assert.Equal("Dave", extractedData[0].Name);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldTransformDataDuringExtraction()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new Person
            {
                Name = row.Cell(1).GetString().ToUpper(),
                Age = !row.Cell(2).IsEmpty() ? (int)(row.Cell(2).GetDouble() * 2) : null,
                Salary = !row.Cell(3).IsEmpty() ? (decimal)(row.Cell(3).GetDouble() / 2) : null,
                JoinDate = row.Cell(4).GetDateTime().AddYears(1),
                LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : DateTime.Now
            });

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
        Assert.Equal("ALICE", extractedData[0].Name);
        Assert.Equal(50, extractedData[0].Age);
        Assert.Equal(25000.375m, extractedData[0].Salary);
        Assert.Equal(new DateTime(2021, 5, 1), extractedData[0].JoinDate);
        Assert.NotNull(extractedData[0].LastActive);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldCreateDifferentObjectType()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new CustomPerson
            {
                FullName = row.Cell(1).GetString(),
                YearsOld = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : 0,
                AnnualSalary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : 0,
                StartDate = row.Cell(4).GetDateTime(),
                IsActive = !row.Cell(5).IsEmpty()
            });

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
        Assert.Equal("Alice", extractedData[0].FullName);
        Assert.Equal(25, extractedData[0].YearsOld);
        Assert.Equal(50000.75m, extractedData[0].AnnualSalary);
        Assert.Equal(new DateTime(2020, 5, 1), extractedData[0].StartDate);
        Assert.False(extractedData[0].IsActive);

        Assert.True(extractedData[1].IsActive);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldHandleEmptyWorksheet()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");
            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new Person
            {
                Name = row.Cell(1).GetString(), Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null
            });

        Assert.NotNull(extractedData);
        Assert.Empty(extractedData);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldSelectSpecificColumns()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new Person
            {
                Name = row.Cell(1).GetString(), JoinDate = row.Cell(4).GetDateTime()
            });


        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
        Assert.Equal("Alice", extractedData[0].Name);
        Assert.Equal(new DateTime(2020, 5, 1), extractedData[0].JoinDate);
        Assert.Null(extractedData[0].Age);
        Assert.Null(extractedData[0].Salary);
        Assert.Null(extractedData[0].LastActive);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldThrowWhenExtractCalledWithoutDelegate()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream);


        var exception = Record.Exception(() => extractor.ExtractWithManualMapping<Person>(null!));
        Assert.NotNull(exception);
        Assert.IsType<InvalidOperationException>(exception);
        Assert.Equal("A row mapping function must be provided when manual mapping is enabled.", exception.Message);
    }

    [Fact]
    public void ExcelExtractor_WithManualMapping_ShouldIgnoreAttributeMappings()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .ExtractWithManualMapping(row => new NoExcelAttributes
            {
                Name = row.Cell(1).GetString(),
                Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
                Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
                JoinDate = row.Cell(4).GetDateTime(),
                LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
            });

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);
        Assert.Equal("Alice", extractedData[0].Name);
        Assert.Equal(25, extractedData[0].Age);
    }

    [Fact]
    public void ExcelExtractor_ShouldGetTimeSpanValue()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Full Name";
            worksheet.Cell(1, 2).Value = "Work Start Time";

            worksheet.Cell(2, 1).Value = "Alice";
            worksheet.Cell(2, 2).Value = new TimeSpan(9, 0, 0);

            worksheet.Cell(3, 1).Value = "Bob";
            worksheet.Cell(3, 2).Value = new TimeSpan(13, 30, 0);

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonWithTimeOnly>();

        Assert.NotNull(extractedData);
        Assert.Equal(2, extractedData.Count);
        Assert.Equal(new TimeSpan(9, 0, 0), extractedData[0].WorkStartTime);
        Assert.Equal(new TimeSpan(13, 30, 0), extractedData[1].WorkStartTime);
    }


    [Fact]
    public void ExcelExtractor_ShouldThrowExceptionWhenWorksheetIndexIsInvalid()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(999)
            .FromStream(stream);

        var exception = Record.Exception(() => extractor.Extract<Person>());

        Assert.NotNull(exception);
        Assert.IsType<ArgumentOutOfRangeException>(exception);
    }

    [Fact]
    public void ExcelExtractor_ShouldThrowExceptionWhenInvalidFileIsProvided()
    {
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromFile("invalid.xlsx");

        var exception = Record.Exception(() => extractor.Extract<Person>());

        Assert.NotNull(exception);
        Assert.IsType<FileNotFoundException>(exception);
    }

    [Fact]
    public void ExcelExtractor_WithSpecificColumns_ShouldReturnValidData()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonWithSpecificColumns>();

        Assert.NotNull(extractedData);
        Assert.Equal("Alice", extractedData[0].NameOnly);
        Assert.Equal(50000.75m, extractedData[0].SalaryOnly);
        Assert.Equal("Charlie", extractedData[2].NameOnly);
        Assert.Equal(-100.50m, extractedData[2].SalaryOnly);
    }

    [Fact]
    public void ExcelExtractor_ShouldParseGuidPropertiesCorrectly()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "UserId";

            worksheet.Cell(2, 1).Value = "Alice";
            worksheet.Cell(2, 2).Value = "123e4567-e89b-12d3-a456-426614174000";

            worksheet.Cell(3, 1).Value = "Bob";
            worksheet.Cell(3, 2).Value = "00000000-0000-0000-0000-000000000000";

            worksheet.Cell(4, 1).Value = "Charlie";
            worksheet.Cell(4, 2).Clear();
            worksheet.Cell(4, 3).Clear();

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractedData = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonWithGuidAndEnum>();

        Assert.NotNull(extractedData);
        Assert.Equal(3, extractedData.Count);

        Assert.Equal("Alice", extractedData[0].Name);
        Assert.Equal(Guid.Parse("123e4567-e89b-12d3-a456-426614174000"), extractedData[0].UserId);

        Assert.Equal("Bob", extractedData[1].Name);
        Assert.Equal(Guid.Empty, extractedData[1].UserId);

        Assert.Equal("Charlie", extractedData[2].Name);

    }

    
    [Fact]
    public void ExcelExtractor_WithoutHeader_ShouldParseGuid()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Alice";
            worksheet.Cell(1, 2).Value = "123e4567-e89b-12d3-a456-426614174000";

            worksheet.Cell(2, 1).Value = "Bob";
            worksheet.Cell(2, 2).Value = "00000000-0000-0000-0000-000000000000";

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractedData = new ExcelExtractor()
            .WithHeader(false)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonNoHeaderWithGuidAndEnum>();

        Assert.NotNull(extractedData);
        Assert.Equal(2, extractedData.Count);

        Assert.Equal("Alice", extractedData[0].Name);
        Assert.Equal(Guid.Parse("123e4567-e89b-12d3-a456-426614174000"), extractedData[0].UserId);

        Assert.Equal("Bob", extractedData[1].Name);
        Assert.Equal(Guid.Empty, extractedData[1].UserId);
    }
    
    [Fact]
    public void ExcelExtractor_ShouldHandleInvalidGuid()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "UserId";

            worksheet.Cell(2, 1).Value = "Alice";
            worksheet.Cell(2, 2).Value = "not-a-valid-guid";

            worksheet.Cell(3, 1).Value = "Bob";
            worksheet.Cell(3, 2).Clear();

            workbook.SaveAs(stream);
        }
        stream.Position = 0;

        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream);
           

        var exception = Record.Exception(() => extractor.Extract<PersonWithGuidAndEnum>());

        Assert.NotNull(exception);
        Assert.IsType<InvalidOperationException>(exception);
    }
}