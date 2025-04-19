using XtractXcel;
namespace ExcelTransformLoad.UnitTests.TestHelpers;

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

public class PersonWithSpecificColumns
{
    [ExcelColumn("Full Name", "Name", "Employee Name")]
    public string? NameOnly { get; init; }

    [ExcelColumn("Salary")]
    public decimal? SalaryOnly { get; init; }
}

public class NoExcelAttributes
{
    public string? Name { get; init; }
    public int? Age { get; init; }
    public decimal? Salary { get; init; }
    public DateTime JoinDate { get; init; }
    public DateTime? LastActive { get; init; }
}

public class PersonNoHeader
{
    public string? Name { get; init; }
    public int? Age { get; init; }
    public decimal? Salary { get; init; }
    public DateTime JoinDate { get; init; }
    public DateTime? LastActive { get; init; }
}

public class CustomPerson
{
    public string FullName { get; init; } = string.Empty;
    public int YearsOld { get; init; }
    public decimal AnnualSalary { get; init; }
    public DateTime StartDate { get; init; }
    public bool IsActive { get; init; }
}

public class PersonWithTimeOnly
{
    [ExcelColumn("Full Name")]
    public string? Name { get; init; }

    [ExcelColumn("Work Start Time")]
    public TimeSpan WorkStartTime { get; init; }
}

public class PersonWithGuidAndEnum
{
    [ExcelColumn("Name")]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("UserId")]
    public Guid UserId { get; set; }
}

public class PersonNoHeaderWithGuidAndEnum
{
    public string Name { get; set; } = string.Empty;

    public Guid? UserId { get; set; }
}