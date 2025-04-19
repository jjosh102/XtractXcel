using XtractXcel;

namespace ExcelTransformLoad.Benchmarks;

public class Person
{
    [ExcelColumn("Full Name", "Name", "Employee Name")]
    public string? Name { get; init; }

    [ExcelColumn("Age", "Employee Age")]
    public int? Age { get; init; }

    [ExcelColumn("Salary")]
    public decimal? Salary { get; init; }

    [ExcelColumn("Join Date")]
    public DateTime JoinDate { get; init; }

    [ExcelColumn("Last Active", "Last Activity")]
    public DateTime? LastActive { get; init; }
}
