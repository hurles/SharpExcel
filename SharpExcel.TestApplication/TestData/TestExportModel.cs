using System.ComponentModel.DataAnnotations;
using SharpExcel.Models.Attributes;

namespace SharpExcel.TestApplication.TestData;

public class TestExportModel
{
    [SharpExcelColumnDefinition(columnName: "ID", width: 20)]
    public int Id { get; set; }
    
    [SharpExcelColumnDefinition(columnName: "Status", width: 15)]
    public TestStatus Status { get; set; }

    [StringLength(10)]
    [SharpExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string? FirstName { get; set; } = null!;

    [StringLength(20)]
    [SharpExcelColumnDefinition(columnName: "Last Name", width: 25)]
    public string? LastName { get; set; } = null!;
    
    [Required]
    [SharpExcelColumnDefinition(columnName: "Email", width: 50)]
    public string? Email { get; set; } = null!;
    
    [SharpExcelColumnDefinition(columnName: "Budget", width: 15)]
    public decimal Budget { get; set; }
    
    [SharpExcelColumnDefinition(columnName: "Department", width: 15)]
    public TestDepartment TestDepartment { get; set; }
}