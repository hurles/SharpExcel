using System.ComponentModel.DataAnnotations;

namespace SharpExcel.Tests.Shared;

public enum TestEnum
{
    [Display(Name = "DisplayValue")]
    ValueA,
    ValueB,
    ValueC
}