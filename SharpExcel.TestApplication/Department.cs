using System.ComponentModel.DataAnnotations;

namespace SharpExcel.TestApplication;

public enum Department
{
    Unknown,
    [Display(Name = "Accounting")]
    ValueA,
    [Display(Name = "Finance")]
    ValueB,
    [Display(Name = "HR")]
    ValueC
}