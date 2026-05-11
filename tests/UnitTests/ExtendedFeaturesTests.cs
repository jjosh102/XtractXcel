using ExcelTransformLoad.UnitTests.TestHelpers;
using XtractXcel;

namespace ExcelTransformLoad.UnitTests;

public class ExtendedFeaturesTests
{
    [Fact]
    public void Extract_WithRowAction_ShouldModifyData()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithWorksheetIndex(1)
            .FromStream(stream);

        var data = extractor.Extract<PersonWithGuidAndEnum>(p => p.Name = p.Name?.ToUpper() ?? string.Empty);

        Assert.All(data, p => Assert.Equal(p.Name?.ToUpper(), p.Name));
    }

    [Fact]
    public void RegisterConverter_ShouldUseCustomConversion()
    {
        try
        {
            ExcelConverterRegistry.RegisterConverter<int>((prop) =>
            {
                if (prop.Name == "Age")
                {
                    return (obj, val) =>
                    {
                        if (val == null)
                        {
                            prop.SetValue(obj, null);
                            return;
                        }

                        int age = Convert.ToInt32(val);
                        prop.SetValue(obj, age == 25 ? 100 : age);
                    };
                }

                return (obj, val) => prop.SetValue(obj, val == null ? null : Convert.ToInt32(val));
            });

            using var stream = TestExcelGenerator.CreateTestExcelFile();
            var extractor = new ExcelExtractor()
                .WithHeader(true)
                .WithWorksheetIndex(1)
                .FromStream(stream);

            var data = extractor.Extract<PersonWithCustomConverter>();

            Assert.Equal(100, data[0].Age); 
            Assert.NotEqual(100, data[2].Age); 
        }
        finally
        {
            ExcelConverterRegistry.Clear();
            ExcelExtractor.ClearCache();
        }
    }

    [Fact]
    public async Task ExtractAsync_ShouldReturnDataCorrectly()
    {
        using var stream = TestExcelGenerator.CreateTestExcelFile();
        var extractor = new ExcelExtractor()
            .WithHeader(true)
            .WithWorksheetIndex(1)
            .FromStream(stream);

        var data = await extractor.ExtractAsync<Person>();

        Assert.NotNull(data);
        Assert.Equal(3, data.Count);
        Assert.Equal("Alice", data[0].Name);
    }

    [Fact]
    public void ExcelExtractor_ShouldNotLockFileAfterExtraction()
    {
        var tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
        try
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cell(1, 1).Value = "Name";
                ws.Cell(2, 1).Value = "Test";
                workbook.SaveAs(tempFile);
            }

            var extractor = new ExcelExtractor()
                .WithWorksheetIndex(1)
                .WithHeader(true)
                .FromFile(tempFile);

            var data = extractor.Extract<PersonWithAttribute>();
            Assert.Single(data);
            File.Delete(tempFile);
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                try
                {
                    File.Delete(tempFile);
                }
                catch
                {
                }
            }
        }
    }

    public class PersonWithCustomConverter
    {
        [ExcelColumn("Full Name", "Name", "Employee Name")]
        public string? Name { get; init; }

        [ExcelColumn("Age", "Employee Age")]
        public int? Age { get; init; }
    }

    public class PersonWithAttribute
    {
        [ExcelColumn("Name")]
        public string? Name { get; set; }
    }
}
