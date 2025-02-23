
namespace ExcelTransformLoad;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    public string ColumnName { get; }
    public ExcelColumnAttribute(string columnName) => ColumnName = columnName;
}
