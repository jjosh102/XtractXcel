namespace ExcelTransformLoad.Extractor;

public record ExcelDataSourceOptions
{
    public string? FilePath { get; init; }
    public Stream? Stream { get; init; }
    public SourceType Source =>
        !string.IsNullOrEmpty(FilePath) ? SourceType.FilePath :
        Stream != null ? SourceType.Stream :
        SourceType.None;
}