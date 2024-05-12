using SharpExcel.Models.Attributes;

namespace SharpExcel.Tests.Shared;

public class TestModel
{
    [ExcelColumnDefinition(columnName: "ID", width: 45)]
    public int Id { get; set; }

    [ExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string FirstName { get; set; } = null!;

    [ExcelColumnDefinition(columnName: "Last Name", width: 50)]
    public string LastName { get; set; } = null!;

    [ExcelColumnDefinition(columnName: "Test value", width: 50)]
    public TestEnum TestValue { get; set; } = TestEnum.ValueA;
}