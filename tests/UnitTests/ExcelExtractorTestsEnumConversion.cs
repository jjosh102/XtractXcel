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

            worksheet.Cell(5, 1).Value = "Josh";
            worksheet.Cell(5, 2).Value = "Default";

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithWorksheetIndex(1)
            .FromStream(stream)
            .Extract<PersonWithEnumStatus>();

        Assert.NotNull(extractor);
        Assert.Equal(4, extractor.Count);

        Assert.Equal("Alice", extractor[0].Name);
        Assert.Equal(UserStatus.Active, extractor[0].Status);

        Assert.Equal("Bob", extractor[1].Name);
        Assert.Equal(UserStatus.Inactive, extractor[1].Status);

        Assert.Equal("Charlie", extractor[2].Name);
        Assert.Equal(UserStatus.Suspended, extractor[2].Status);

        Assert.Equal("Josh", extractor[3].Name);
        Assert.Equal(UserStatus.None, extractor[3].Status);
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

            worksheet.Cell(3, 1).Value = "Josh";
            worksheet.Cell(3, 2).Value = "Default";

            workbook.SaveAs(stream);
        }

        stream.Position = 0;

        var extractor = new ExcelExtractor()
            .WithHeader(false)
            .WithWorksheetIndex(1)
            .FromStream(stream)
            .Extract<PersonNoHeaderWithEnumStatus>();

        Assert.NotNull(extractor);
        Assert.Equal(3, extractor.Count);

        Assert.Equal("Alice", extractor[0].Name);
        Assert.Equal(UserStatus.Active, extractor[0].Status);

        Assert.Equal("Bob", extractor[1].Name);
        Assert.Equal(UserStatus.Inactive, extractor[1].Status);

        Assert.Equal("Josh", extractor[2].Name);
        Assert.Equal(UserStatus.None, extractor[2].Status);
    }


    public enum UserStatus
    {
        None,
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