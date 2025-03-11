
using ExcelTransformLoad.Extractor;
using ExcelTransformLoad.UnitTests.TestHelpers;

namespace ExcelTransformLoad.UnitTests;


public class ExcelExtacrtorTestsEnsure
{
    [Fact]
    public void ExcelExtractor_EnsureSourceNotSet_ShouldThrowIfSourceIsSet()
    {
        var extractor = new ExcelExtractor()
            .FromStream(new MemoryStream());

        var exception = Assert.Throws<InvalidOperationException>(() => extractor.WithHeader(true));
        Assert.Equal("Source (file or stream) has already been set. Cannot modify settings after source is set.", exception.Message);
    }

    [Fact]
    public void ExcelExtractor_EnsureSourceIsSet_ShouldThrowIfSourceIsNotSet()
    {
        var extractor = new ExcelExtractor();
        var exception = Assert.Throws<InvalidOperationException>(() => extractor.Extract<Person>());
        Assert.Equal("Data source (file or stream) is required before extraction.", exception.Message);
    }

    // [Fact]
    // public void ExcelExtractor_EnsureNotExtractingMultiple_ShouldThrowIfExtractingMultiple()
    // {
    //     var exception = Assert.Throws<InvalidOperationException>(() =>
    //         new ExcelExtractor()
    //             .WitMultipleWorksheet()
    //             .WithWorksheetIndex(1)
    //             .FromStream(new MemoryStream()).ExtractMultiple());
    //     Assert.Equal("Cannot use WithWorksheetIndex or WithSheetName when extracting multiple worksheets.", exception.Message);
    // }

    [Fact]
    public void ExcelExtractor_EnsureWorksheetIndexNotSet_ShouldThrowIfWorksheetIndexIsSet()
    {
        var extractor = new ExcelExtractor()
            .WithWorksheetIndex(1);

        var exception = Assert.Throws<InvalidOperationException>(() => extractor.WithWorksheetName("Sheet1"));
        Assert.Equal("Worksheet index is already set. Cannot set worksheet name when index is already specified.", exception.Message);
    }

    [Fact]
    public void ExcelExtractor_EnsureWorksheetNameNotSet_ShouldThrowIfWorksheetNameIsSet()
    {
        var extractor = new ExcelExtractor()
            .WithWorksheetName("Sheet1");

        var exception = Assert.Throws<InvalidOperationException>(() => extractor.WithWorksheetIndex(1));
        Assert.Equal("Worksheet name is already set. Cannot set worksheet index when name is already specified.", exception.Message);
    }

    [Fact]
    public void ExcelExtractor_EnsureNotToUseExtractWhenExtractingMultiple_ShouldThrowIfExtractingMultiple()
    {
        var extractor = new ExcelExtractor()
            .WitMultipleWorksheet()
            .FromStream(new MemoryStream());

        var exception = Assert.Throws<InvalidOperationException>(() => extractor.Extract<Person>());
        Assert.Equal("Cannot use Extract. Use ExtractMultiple when WithMultipleWorksheet is enabled.", exception.Message);
    }
}