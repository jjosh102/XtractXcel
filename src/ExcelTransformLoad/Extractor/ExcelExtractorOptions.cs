namespace ExcelTransformLoad.Extractor;

public class ExcelExtractorOptions
{
    public bool ReadHeader { get; set; } = true;
    public int SheetIndex { get; set; } = 1;
    
    private string? _filePath;
    private Stream? _stream;
    
    public SourceType Source { get; private set; } = SourceType.None;
    
    public string? FilePath
    {
        get => _filePath;
        set
        {
            if (!string.IsNullOrEmpty(value))
            {
                _filePath = value;
                _stream = null;
                Source = SourceType.FilePath;
            }
            else
            {
                _filePath = null;
                if (_stream == null)
                    Source = SourceType.None;
            }
        }
    }
    
    public Stream? Stream
    {
        get => _stream;
        set
        {
            if (value != null)
            {
                _stream = value;
                _filePath = null;
                Source = SourceType.Stream;
            }
            else
            {
                _stream = null;
                if (_filePath == null)
                    Source = SourceType.None;
            }
        }
    }
    
}