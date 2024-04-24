using System.ComponentModel.DataAnnotations;

namespace SharpExcel.Models.Results;

public struct ExcelCellValidationResult
{
    public ExcelAddress Address { get; set; }
    
    public List<ValidationResult> ValidationResults { get; set; } 
}