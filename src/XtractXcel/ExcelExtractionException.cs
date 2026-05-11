namespace XtractXcel;

public class ExcelExtractionException : Exception
{
    public int RowNumber { get; }
    public int ColumnIndex { get; }
    public string? ColumnName { get; }
    public object? Value { get; }

    public ExcelExtractionException(string message, int rowNumber, int columnIndex, string? columnName, object? value, Exception? innerException = null)
        : base(message, innerException)
    {
        RowNumber = rowNumber;
        ColumnIndex = columnIndex;
        ColumnName = columnName;
        Value = value;
    }
}
