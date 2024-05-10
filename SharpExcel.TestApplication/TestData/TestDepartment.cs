using System.ComponentModel.DataAnnotations;

namespace SharpExcel.TestApplication.TestData;

public enum TestDepartment
{
    Unknown,
    [Display(Name = "Accounting")]
    ValueA,
    [Display(Name = "Finance")]
    ValueB,
    [Display(Name = "HR")]
    ValueC
}