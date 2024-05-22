using SharpExcel.Models.Attributes;

namespace SharpExcel.Tests.Shared;

public class TestModel
{
    [SharpExcelColumnDefinition(columnName: "ID", width: 45)]
    public int Id { get; set; }

    [SharpExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string FirstName { get; set; } = null!;

    [SharpExcelColumnDefinition(columnName: "Last Name", width: 50)]
    public string LastName { get; set; } = null!;

    [SharpExcelColumnDefinition(columnName: "Test value", width: 50)]
    public TestEnum TestValue { get; set; } = TestEnum.ValueA;
}