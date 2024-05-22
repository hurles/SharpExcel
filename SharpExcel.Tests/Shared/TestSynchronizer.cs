using Microsoft.Extensions.Options;
using SharpExcel.Exporters;
using SharpExcel.Models.Configuration;

namespace SharpExcel.Tests.Shared;

public class TestSynchronizer : BaseSharpExcelSynchronizer<TestModel>
{
    public TestSynchronizer(IOptions<ExporterOptions<TestModel>> options) : base(options)
    {
    }
}