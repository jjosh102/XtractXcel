using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public sealed class ExcelExtractor
{
    private bool _isSourceSet = false;
    private string? _filPath;
    private Stream? _stream;
    private bool _readHeader = false;
    private int _sheetIndex = 1;


    public ExcelExtractor WithHeader(bool readHeader)
    {
        EnsureSourceNotSet();
        _readHeader = readHeader;
        return this;
    }

    public ExcelExtractor WithSheetIndex(int sheetIndex)
    {
        EnsureSourceNotSet();
        _sheetIndex = sheetIndex;
        return this;
    }

    public ExcelExtractor FromFile(string filePath)
    {
        EnsureSourceNotSet();

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        }

        _filPath = filePath;
        _isSourceSet = true;
        return this;
    }

    public ExcelExtractor FromStream(Stream stream)
    {
        EnsureSourceNotSet();

        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream), "Stream cannot be null");
        }

        _stream = stream;
        _isSourceSet = true;
        return this;
    }

    public List<T> Extract<T>() where T : new()
    {
        if (!_isSourceSet)
        {
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
        }

        var _options = new ExcelDataSourceOptions
        {
            FilePath = _filPath,
            Stream = _stream,
        };

        var extractor = new ExcelDataExtractor(_options);
        return extractor.ExtractData<T>(_sheetIndex, _readHeader);
    }

    public List<T> ExtractWithManualMapping<T>(Func<IXLRangeRow, T> manualMapping) where T : new()
    {
        if (!_isSourceSet)
        {
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
        }

        var _options = new ExcelDataSourceOptions
        {
            FilePath = _filPath,
            Stream = _stream,
        };

        var extractor = new ExcelDataExtractor(_options);

        if (manualMapping is null)
        {
            throw new InvalidOperationException("A row mapping function must be provided when manual mapping is enabled.");
        }

        return extractor.ExtractData<T>(_sheetIndex, manualMapping, _readHeader);
    }

    private void EnsureSourceNotSet()
    {
        if (_isSourceSet)
        {
            throw new InvalidOperationException("Source (file or stream) has already been set. Cannot modify settings after source is set.");
        }
    }
}
