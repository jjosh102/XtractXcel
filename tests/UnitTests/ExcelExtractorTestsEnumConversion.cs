using ClosedXML.Excel;
using ExcelTransformLoad.Extractor;

namespace ExcelTransformLoad.UnitTests;

public class ExcelExtractorTestsEnumConversion
{
    [Fact]
    public void ExcelExtractor_ShouldParseEnumPropertiesCorrectly()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Status";

            worksheet.Cell(2, 1).Value = "Alice";
            worksheet.Cell(2, 2).Value = "Active";

            worksheet.Cell(3, 1).Value = "Bob";
            worksheet.Cell(3, 2).Value = "Inactive";

            worksheet.Cell(4, 1).Value = "Charlie";
            worksheet.Cell(4, 2).Value = "Suspended";

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonWithEnumStatus>();

        Assert.NotNull(extractor);
        Assert.Equal(3, extractor.Count);

        Assert.Equal("Alice", extractor[0].Name);
        Assert.Equal(UserStatus.Active, extractor[0].Status);

        Assert.Equal("Bob", extractor[1].Name);
        Assert.Equal(UserStatus.Inactive, extractor[1].Status);

        Assert.Equal("Charlie", extractor[2].Name);
        Assert.Equal(UserStatus.Suspended, extractor[2].Status);
    }

    [Fact]
    public void ExcelExtractor_ShouldHandleInvalidEnumValues()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Status";

            worksheet.Cell(2, 1).Value = "Alice";
            worksheet.Cell(2, 2).Value = "Unknown";

            worksheet.Cell(3, 1).Value = "Bob";
            worksheet.Cell(3, 2).Clear();

            worksheet.Cell(4, 1).Value = "Charlie";
            worksheet.Cell(4, 2).Value = 999;

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithSheetIndex(1)
            .FromStream(stream);
      
        var exception = Record.Exception(() => extractor.Extract<PersonWithEnumStatus>());
        Assert.NotNull(exception);
        Assert.IsType<InvalidOperationException>(exception);
        Assert.Equal("Error setting value for column 2 in row 2",exception.Message);
    }


    [Fact]
    public void ExcelExtractor_WithoutHeader_ShouldParseEnumProperties()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Alice";
            worksheet.Cell(1, 2).Value = "Active";

            worksheet.Cell(2, 1).Value = "Bob";
            worksheet.Cell(2, 2).Value = "Inactive";

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractor = new ExcelExtractor()
            .WithHeader(false)
            .WithSheetIndex(1)
            .FromStream(stream)
            .Extract<PersonNoHeaderWithEnumStatus>();

        Assert.NotNull(extractor);
        Assert.Equal(2, extractor.Count);

        Assert.Equal("Alice", extractor[0].Name);
        Assert.Equal(UserStatus.Active, extractor[0].Status);

        Assert.Equal("Bob", extractor[1].Name);
        Assert.Equal(UserStatus.Inactive, extractor[1].Status);
    }


    public enum UserStatus
    {
        Active,
        Inactive,
        Suspended
    }


    public class PersonWithEnumStatus
    {
        [ExcelColumn("Name")] public string? Name { get; set; }

        [ExcelColumn("Status")] public UserStatus Status { get; set; }
    }

    public class PersonNoHeaderWithEnumStatus
    {
        public string? Name { get; set; }

        public UserStatus? Status { get; set; }
    }
}