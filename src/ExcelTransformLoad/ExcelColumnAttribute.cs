[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    public string[] ColumnNames { get; }
    public ExcelColumnAttribute(params string[] columnNames) => ColumnNames = columnNames ?? [];
}
