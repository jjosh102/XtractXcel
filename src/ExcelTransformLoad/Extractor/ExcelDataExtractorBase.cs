
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

internal abstract class ExcelDataExtractorBase : IDisposable
{
    private readonly ExcelDataSourceOptions _options;
    private XLWorkbook? _workbook;
    private bool _isDisposed = false;

    protected ExcelDataExtractorBase(ExcelDataSourceOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    protected XLWorkbook GetOrCreateWorkbook()
    {
        if (_isDisposed)
        {
            throw new ObjectDisposedException(nameof(ExcelDataExtractorBase));
        }

        if (_workbook is null)
        {
            _workbook = _options.Source switch
            {
                SourceType.FilePath => new XLWorkbook(_options.FilePath!),
                SourceType.Stream => new XLWorkbook(_options.Stream!),
                SourceType.None => throw new InvalidOperationException("Either a file path or a stream must be provided."),
                _ => throw new InvalidOperationException("Invalid source type.")
            };
        }

        return _workbook;
    }

    public void Dispose()
    {
        if (!_isDisposed)
        {
            _workbook?.Dispose();
            _workbook = null;
            _isDisposed = true;
        }
    }
}