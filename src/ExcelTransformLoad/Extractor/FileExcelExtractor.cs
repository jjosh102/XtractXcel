
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public class FileExcelExtractor<T> : ExcelExtractorBase<T> where T : new()
{
    private readonly string _filePath;

    public FileExcelExtractor(string filePath)
    {
        ArgumentNullException.ThrowIfNullOrWhiteSpace(filePath, nameof(filePath));
        _filePath = filePath;
    }

    protected override XLWorkbook GetWorkbook()
    {
        return new XLWorkbook(_filePath);
    }
}