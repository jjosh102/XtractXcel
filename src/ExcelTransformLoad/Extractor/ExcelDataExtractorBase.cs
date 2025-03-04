
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

internal abstract class ExcelDataExtractorBase : IDisposable
{
    protected readonly ExcelDataSourceOptions _options;
    protected XLWorkbook? _workbook;
    protected bool IsDisposed = false;

    protected ExcelDataExtractorBase(ExcelDataSourceOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    protected XLWorkbook GetOrCreateWorkbook()
    {
        if (IsDisposed)
        {
            throw new ObjectDisposedException(nameof(ExcelDataExtractorBase));
        }

        if (_workbook == null)
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
        if (!IsDisposed)
        {
            _workbook?.Dispose();
            _workbook = null;
            IsDisposed = true;
        }
    }
}