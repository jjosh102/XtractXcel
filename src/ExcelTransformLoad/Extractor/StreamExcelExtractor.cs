using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public class StreamExcelExtractor<T> : ExcelExtractorBase<T> where T : new()
{
    private readonly Stream _stream;

    public StreamExcelExtractor(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream, nameof(stream));
        _stream = stream;
    }

    protected override XLWorkbook GetWorkbook()
    {
        return new XLWorkbook(_stream);
    }
}