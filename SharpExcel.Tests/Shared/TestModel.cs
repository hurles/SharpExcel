using SharpExcel.Abstraction;
using SharpExcel.Attributes;

namespace SharpExcel.Tests.Shared;

public class TestModel : IExcelModel
{
    [ExcelColumnDefinition(columnName: "ID", width: 45)]
    public int Id { get; set; }

    [ExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string FirstName { get; set; } = null!;

    [ExcelColumnDefinition(columnName: "Last Name", width: 50)]
    public string LastName { get; set; } = null!;
}