namespace ExcelTransformLoad.Extractor;

public class ExcelExtractor<T> where T : new()
{
    private readonly ExcelExtractorOptions _options = new();
    private bool _isSourceSet = false;

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

    public IReadOnlyList<T> Extract()
    {
        if (!_isSourceSet)
        {
            throw new InvalidOperationException("Source (file or stream) must be set before extracting data");
        }
        
        var extractor = new ExcelDataExtractor<T>(_options);
        return extractor.ExtractData();
    }

    private void EnsureSourceNotSet()
    {
        if (_isSourceSet)
        {
            throw new InvalidOperationException("Source (file or stream) has already been set. Cannot modify settings after source is set.");
        }
    }
}