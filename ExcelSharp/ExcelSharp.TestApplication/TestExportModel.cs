using ExcelSharp.Attributes;

namespace ExcelSharp.TestApplication;

public class TestExportModel
{
    [ExcelColumnDefinition(columnName: "ID", width: 20)]
    public int Id { get; set; }

    [ExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string? FirstName { get; set; } = null!;

    [ExcelColumnDefinition(columnName: "Last Name", width: 25)]
    public string? LastName { get; set; } = null!;
    
    [ExcelColumnDefinition(columnName: "Email", width: 50)]
    public string? Email { get; set; } = null!;
    
    [ExcelColumnDefinition(columnName: "Budget", width: 15)]
    public decimal Budget { get; set; }
}