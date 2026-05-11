using ClosedXML.Excel;

namespace XtractXcel;

public record ExcelExtractor(
    string? FilePath = null,
    Stream? Stream = null,
    bool ReadHeader = true,
    int? WorksheetIndex = null,
    string? WorksheetName = null
)
{
    public static ExcelExtractor Create()
    {
        return new ExcelExtractor();
    }

    public ExcelExtractor WithHeader(bool readHeader)
    {
        EnsureSourceNotSet();
        return this with { ReadHeader = readHeader };
    }

    public ExcelExtractor WithWorksheetIndex(int workSheetIndex)
    {
        EnsureSourceNotSet();
        return this with { WorksheetIndex = workSheetIndex, WorksheetName = null };
    }

    public ExcelExtractor WithWorksheetName(string workSheetName)
    {
        EnsureSourceNotSet();
        return this with { WorksheetName = workSheetName, WorksheetIndex = null };
    }

    public ExcelExtractor FromFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        }

        EnsureSourceNotSet();
        return this with { FilePath = filePath, Stream = null };
    }

    public ExcelExtractor FromStream(Stream stream)
    {
        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream), "Stream cannot be null");
        }

        EnsureSourceNotSet();
        return this with { Stream = stream, FilePath = null };
    }

    public List<T> Extract<T>(Action<T>? onRowExtracted = null)
        where T : new()
    {
        EnsureSourceIsSet();

        var options = new ExcelDataSourceOptions { FilePath = FilePath, Stream = Stream };
        using var extractor = new ExcelDataExtractor(options);

        if (WorksheetIndex.HasValue)
        {
            return extractor.ExtractData<T>(WorksheetIndex.Value, ReadHeader, onRowExtracted);
        }

        if (!string.IsNullOrEmpty(WorksheetName))
        {
            return extractor.ExtractData<T>(WorksheetName, ReadHeader, onRowExtracted);
        }

        return extractor.ExtractData<T>(1, ReadHeader, onRowExtracted);
    }

    public List<T> ExtractWithManualMapping<T>(Func<IXLRangeRow, T> manualMapping, Action<T>? onRowExtracted = null)
        where T : new()
    {
        EnsureSourceIsSet();

        if (manualMapping is null)
        {
            throw new InvalidOperationException(
                "A row mapping function must be provided when manual mapping is enabled.");
        }

        var options = new ExcelDataSourceOptions { FilePath = FilePath, Stream = Stream };
        using var extractor = new ExcelDataExtractor(options);

        if (WorksheetIndex.HasValue)
        {
            return extractor.ExtractData(WorksheetIndex.Value, manualMapping, ReadHeader, onRowExtracted);
        }

        if (!string.IsNullOrEmpty(WorksheetName))
        {
            return extractor.ExtractData(WorksheetName, manualMapping, ReadHeader, onRowExtracted);
        }

        return extractor.ExtractData(1, manualMapping, ReadHeader, onRowExtracted);
    }

    public async Task<List<T>> ExtractAsync<T>(CancellationToken cancellationToken = default, Action<T>? onRowExtracted = null)
        where T : new()
    {
        EnsureSourceIsSet();

        var options = new ExcelDataSourceOptions { FilePath = FilePath, Stream = Stream };
        using var extractor = new ExcelDataExtractor(options);

        if (WorksheetIndex.HasValue)
        {
            return await extractor.ExtractDataAsync<T>(WorksheetIndex.Value, ReadHeader, cancellationToken, onRowExtracted);
        }

        if (!string.IsNullOrEmpty(WorksheetName))
        {
            return await extractor.ExtractDataAsync<T>(WorksheetName, ReadHeader, cancellationToken, onRowExtracted);
        }

        return await extractor.ExtractDataAsync<T>(1, ReadHeader, cancellationToken, onRowExtracted);
    }

    public async Task<List<T>> ExtractWithManualMappingAsync<T>(Func<IXLRangeRow, T> manualMapping, CancellationToken cancellationToken = default, Action<T>? onRowExtracted = null)
        where T : new()
    {
        EnsureSourceIsSet();

        if (manualMapping is null)
        {
            throw new InvalidOperationException(
                "A row mapping function must be provided when manual mapping is enabled.");
        }

        var options = new ExcelDataSourceOptions { FilePath = FilePath, Stream = Stream };
        using var extractor = new ExcelDataExtractor(options);

        if (WorksheetIndex.HasValue)
        {
            return await extractor.ExtractDataAsync(WorksheetIndex.Value, manualMapping, ReadHeader, cancellationToken, onRowExtracted);
        }

        if (!string.IsNullOrEmpty(WorksheetName))
        {
            return await extractor.ExtractDataAsync(WorksheetName, manualMapping, ReadHeader, cancellationToken, onRowExtracted);
        }

        return await extractor.ExtractDataAsync(1, manualMapping, ReadHeader, cancellationToken, onRowExtracted);
    }

    private void EnsureSourceNotSet()
    {
        if (FilePath is not null || Stream is not null)
        {
            throw new InvalidOperationException(
                "Source (file or stream) has already been set. Cannot modify settings after source is set.");
        }
    }

    public static void ClearCache()
    {
        ExcelObjectExtractorCache.Clear();
    }

    private void EnsureSourceIsSet()
    {
        if (FilePath is null && Stream is null)
        {
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
        }
    }
}
