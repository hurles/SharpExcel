using SharpExcel.Styling;
using SharpExcel.Styling.Colorization;
using SharpExcel.Styling.Text;

namespace SharpExcel.TestApplication;

public class TestExporter : BaseExcelExporter<TestExportModel>
{
    public override SharpExcelCellStyle GetHeaderStyle()
    {
        var headerStyle = SharpExcelCellStyleConstants.DefaultHeaderStyle;
        headerStyle.FontSize = 13.0f;
        headerStyle.TextStyle = TextStyle.Bold;
        headerStyle.BackgroundColor = new SharpExcelColor(200, 200, 200, 255);

        return headerStyle;
    }

    public override SharpExcelCellStyle GetDataStyle(string propertyName, TestExportModel record)
    {
        var dataStyle = SharpExcelCellStyleConstants.DefaultDataStyle;

        if (propertyName == nameof(TestExportModel.Budget))
        {
            switch (record.Budget)
            {
                case < 0:
                    dataStyle.TextColor = SharpExcelColorConstants.Red;
                    break;
                case > 0:
                    dataStyle.TextColor = SharpExcelColorConstants.Green;
                    break;
                default:
                    dataStyle.TextColor = SharpExcelColorConstants.Black;
                    break;
            }
        }

        return dataStyle;
    }
}