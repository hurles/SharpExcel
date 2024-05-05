using System.ComponentModel.DataAnnotations;
using SharpExcel.Models.Attributes;

namespace SharpExcel.TestApplication;

public class TestExportModel
{
    [ExcelColumnDefinition(columnName: "ID", width: 20)]
    public int Id { get; set; }

    [StringLength(10)]
    [ExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string? FirstName { get; set; } = null!;

    [StringLength(20)]
    [ExcelColumnDefinition(columnName: "Last Name", width: 25)]
    public string? LastName { get; set; } = null!;
    
    [Required]
    [ExcelColumnDefinition(columnName: "Email", width: 50)]
    public string? Email { get; set; } = null!;
    
    [ExcelColumnDefinition(columnName: "Budget", width: 15)]
    public decimal Budget { get; set; }
    
    [ExcelColumnDefinition(columnName: "Department", width: 15)]
    public Department Department { get; set; }
}