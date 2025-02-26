using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public class ExcelExtractor<T> where T : new()
{
    private readonly ExcelExtractorOptions _options = new();
    private Func<IXLRangeRow, T>? RowMappingDelegate { get; set; }
    private bool _isSourceSet = false;
    private bool _isManualMappingSet = false;

    public ExcelExtractor<T> WithHeader(bool readHeader)
    {
        EnsureSourceNotSet();
        _options.ReadHeader = readHeader;
        return this;
    }

    public ExcelExtractor<T> WithSheetIndex(int sheetIndex)
    {
        EnsureSourceNotSet();
        _options.SheetIndex = sheetIndex;
        return this;
    }

    public ExcelExtractor<T> WithManualMapping(Func<IXLRangeRow, T> defineRowMapping)
    {
        EnsureSourceNotSet();
        RowMappingDelegate = defineRowMapping;
        _isManualMappingSet = true;
        return this;
    }

    public ExcelExtractor<T> FromFile(string filePath)
    {
        EnsureSourceNotSet();

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        }

        _options.FilePath = filePath;
        _isSourceSet = true;
        return this;
    }

    public ExcelExtractor<T> FromStream(Stream stream)
    {
        EnsureSourceNotSet();

        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream), "Stream cannot be null");
        }

        _options.Stream = stream;
        _isSourceSet = true;
        return this;
    }

    public List<T> Extract()
    {
        if (!_isSourceSet)
        {
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
        }

        var extractor = new ExcelDataExtractor<T>(_options);

        if (_isManualMappingSet && RowMappingDelegate is null)
        {
            throw new InvalidOperationException("A row mapping function must be provided when manual mapping is enabled.");
        }

        return _isManualMappingSet ? extractor.ExtractData(RowMappingDelegate!) : extractor.ExtractData();
    }


    private void EnsureSourceNotSet()
    {
        if (_isSourceSet)
        {
            throw new InvalidOperationException("Source (file or stream) has already been set. Cannot modify settings after source is set.");
        }
    }
}