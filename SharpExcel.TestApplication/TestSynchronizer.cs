using Microsoft.Extensions.Options;
using SharpExcel.Exporters;
using SharpExcel.Models.Configuration;
using SharpExcel.TestApplication.TestData;

namespace SharpExcel.TestApplication;

public class TestSynchronizer : BaseSharpExcelSynchronizer<TestExportModel>
{
    public TestSynchronizer(IOptions<ExporterOptions<TestExportModel>> options) : base(options)
    {
    }
}