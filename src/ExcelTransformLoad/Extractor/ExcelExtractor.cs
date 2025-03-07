using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public sealed class ExcelExtractor
{
    private bool _isSourceSet;
    private string? _filePath;
    private Stream? _stream;
    private bool _readHeader;
    private int _sheetIndex = 1;

    private readonly List<string> _worksheetNames = [];
    private readonly List<int> _worksheetIndexes = [];

    public ExcelExtractor WithHeader(bool readHeader)
    {
        EnsureSourceNotSet();
        _readHeader = readHeader;
        return this;
    }
    public ExcelExtractor WithSheetIndex(int sheetIndex)
    {
        EnsureSourceNotSet();
        if (sheetIndex < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Sheet index must be greater than or equal to 1.");
        }
        _sheetIndex = sheetIndex;
        return this;
    }

    public ExcelExtractor WithSheetIndex(params int[] sheetIndexes)
    {
        EnsureSourceNotSet();

        if (sheetIndexes.Any(i => i < 1))
            throw new ArgumentOutOfRangeException(nameof(sheetIndexes), "Sheet indexes must be >= 1.");

        _worksheetIndexes.AddRange(sheetIndexes);
        return this;
    }

    public ExcelExtractor WithSheetName(params string[] sheetNames)
    {
        EnsureSourceNotSet();
        if (sheetNames.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("Sheet names cannot be null or empty.", nameof(sheetNames));

        _worksheetNames.AddRange(sheetNames);
        return this;
    }

    public ExcelExtractor FromFile(string filePath)
    {
        EnsureSourceNotSet();

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        }


        _filePath = filePath;
        _isSourceSet = true;
        return this;
    }

    public ExcelExtractor FromStream(Stream stream)
    {
        EnsureSourceNotSet();

        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream), "Stream cannot be null");
        }

        _stream = stream;
        _isSourceSet = true;
        return this;
    }

    public List<T> Extract<T>() where T : new()
    {
        EnsureSourceIsSet();

        var options = new ExcelDataSourceOptions
        {
            FilePath = _filePath,
            Stream = _stream,
        };

        return new ExcelDataExtractor(options).ExtractData<T>(_sheetIndex, _readHeader);
    }

    public List<T> ExtractWithManualMapping<T>(Func<IXLRangeRow, T> manualMapping) where T : new()
    {
        EnsureSourceIsSet();

        if (manualMapping is null)
        {
            throw new InvalidOperationException("A row mapping function must be provided when manual mapping is enabled.");
        }

        var options = new ExcelDataSourceOptions
        {
            FilePath = _filePath,
            Stream = _stream,
        };

        return new ExcelDataExtractor(options).ExtractData(_sheetIndex, manualMapping, _readHeader);
    }

    //todo: Support multiple worksheets extraction
    public Dictionary<string, List<object>> ExtractMultiple()
    {
        return [];
    }


    private void EnsureSourceNotSet()
    {
        if (_isSourceSet)
        {
            throw new InvalidOperationException("Source (file or stream) has already been set. Cannot modify settings after source is set.");
        }
    }

    private void EnsureSourceIsSet()
    {
        if (!_isSourceSet)
        {
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
        }
    }
}
