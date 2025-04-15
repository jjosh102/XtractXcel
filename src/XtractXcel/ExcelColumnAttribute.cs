
namespace XtractXcel;

[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnAttribute : Attribute
{
    public string[] ColumnNames { get; }
    public ExcelColumnAttribute(params string[] columnNames) => ColumnNames = columnNames ?? [];
}
