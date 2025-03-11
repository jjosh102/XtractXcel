using System.Reflection;
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public sealed class ExcelExtractor
{
    private bool _isSourceSet;
    private bool _isMultipleExtraction;
    private string? _filePath;
    private Stream? _stream;
    private bool _readHeader;
    private int _workSheetIndex = 1;
    private string _workSheetName = "Sheet1";
    private bool _isWorksheetIndexSet;
    private bool _isWorksheetNameSet;

    public ExcelExtractor WithHeader(bool readHeader)
    {
        EnsureSourceNotSet();
        _readHeader = readHeader;
        return this;
    }

    public ExcelExtractor WitMultipleWorksheet()
    {
        EnsureSourceNotSet();
        EnsureWorksheetNameNotSet();

        _isMultipleExtraction = true;
        return this;
    }

    public ExcelExtractor WithWorksheetIndex(int workSheetIndex)
    {
        EnsureSourceNotSet();
        EnsureNotExtractingMultiple();
        EnsureWorksheetNameNotSet();

        _workSheetIndex = workSheetIndex;
        _isWorksheetIndexSet = true;
        return this;
    }

    public ExcelExtractor WithWorksheetName(string workSheetName)
    {
        EnsureSourceNotSet();
        EnsureNotExtractingMultiple();
        EnsureWorksheetIndexNotSet();

        _workSheetName = workSheetName;
        _isWorksheetNameSet = true;
        return this;
    }

    public ExcelExtractor FromFile(string filePath)
    {
        EnsureSourceNotSet();

        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        _filePath = filePath;
        _isSourceSet = true;
        return this;
    }

    public ExcelExtractor FromStream(Stream stream)
    {
        EnsureSourceNotSet();

        _stream = stream ?? throw new ArgumentNullException(nameof(stream), "Stream cannot be null");
        _isSourceSet = true;
        return this;
    }

    public List<T> Extract<T>() where T : new()
    {
        EnsureSourceIsSet();
        EnsureNotToUseExtractWhenExtractingMultiple();
        var options = new ExcelDataSourceOptions { FilePath = _filePath, Stream = _stream };

        if (_isWorksheetIndexSet)
        {
            return new ExcelDataExtractor(options).ExtractData<T>(_workSheetIndex, _readHeader);
        }

        if (_isWorksheetNameSet)
        {
            return new ExcelDataExtractor(options).ExtractData<T>(_workSheetName, _readHeader);
        }

        return new ExcelDataExtractor(options).ExtractData<T>(1, _readHeader);
    }

    public List<T> ExtractWithManualMapping<T>(Func<IXLRangeRow, T> manualMapping) where T : new()
    {
        EnsureSourceIsSet();
        EnsureNotToUseExtractWhenExtractingMultiple();
        if (manualMapping is null)
        {
            throw new InvalidOperationException(
                "A row mapping function must be provided when manual mapping is enabled.");
        }

        var options = new ExcelDataSourceOptions { FilePath = _filePath, Stream = _stream, };

        if (_isWorksheetIndexSet)
        {
            return new ExcelDataExtractor(options).ExtractData(_workSheetIndex, manualMapping, _readHeader);
        }

        if (_isWorksheetNameSet)
        {
            return new ExcelDataExtractor(options).ExtractData(_workSheetName, manualMapping, _readHeader);
        }

        return new ExcelDataExtractor(options).ExtractData(1, manualMapping, _readHeader);

    }

    // todo: Add support for multiple worksheets
    public ExcelExtractor ExtractWorksheet<T>(string worksheetName, out List<T> data) where T : new()
    {
        EnsureSourceIsSet();
        EnsureNotExtractingMultiple();

        var options = new ExcelDataSourceOptions { FilePath = _filePath, Stream = _stream };
        data = new ExcelDataExtractor(options).ExtractData<T>(worksheetName, _readHeader);

        return this;
    }

    private void EnsureSourceNotSet()
    {
        if (_isSourceSet)
            throw new InvalidOperationException(
                "Source (file or stream) has already been set. Cannot modify settings after source is set.");
    }

    private void EnsureSourceIsSet()
    {
        if (!_isSourceSet)
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
    }

    private void EnsureNotToUseExtractWhenExtractingMultiple()
    {
        if (_isMultipleExtraction)
            throw new InvalidOperationException(
                "Cannot use Extract. Use ExtractMultiple when WithMultipleWorksheet is enabled.");
    }

    private void EnsureNotExtractingMultiple()
    {
        if (_isMultipleExtraction)
            throw new InvalidOperationException(
                "Cannot use WithWorksheetIndex or WithSheetName when extracting multiple worksheets.");
    }

    private void EnsureWorksheetIndexNotSet()
    {
        if (_isWorksheetIndexSet)
            throw new InvalidOperationException(
                "Worksheet index is already set. Cannot set worksheet name when index is already specified.");
    }

    private void EnsureWorksheetNameNotSet()
    {
        if (_isWorksheetNameSet)
            throw new InvalidOperationException(
                "Worksheet name is already set. Cannot set worksheet index when name is already specified.");
    }
}