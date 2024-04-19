using ExcelSharp.Styling;
using ExcelSharp.Styling.Colorization;

namespace ExcelSharp.TestApplication;

public class TestExporter : BaseExcelExporter<TestExportModel>
{
    public override ExcelSharpCellStyle GetHeaderStyle()
    {
        var headerStyle = ExcelSharpCellStyleConstants.DefaultHeaderStyle;
        
        headerStyle.BackgroundColor = new ExcelSharpColor(200, 50, 50, 255);

        return headerStyle;
    }

    public override ExcelSharpCellStyle GetDataStyle(string propertyName, TestExportModel record)
    {
        var dataStyle = ExcelSharpCellStyleConstants.DefaultDataStyle;

        switch (record.Id)
        {
            case 0:
                dataStyle.TextColor = ExcelSharpColorConstants.Red;
                break;
            case 1:
                dataStyle.TextColor = ExcelSharpColorConstants.Yellow;
                break;
            case 2:
                dataStyle.TextColor = ExcelSharpColorConstants.Green;
                break;
            case 3:
                dataStyle.TextColor = ExcelSharpColorConstants.Blue;
                break;
            case 4:
                dataStyle.TextColor = ExcelSharpColorConstants.Purple;
                break;

        }

        return dataStyle;
    }
}