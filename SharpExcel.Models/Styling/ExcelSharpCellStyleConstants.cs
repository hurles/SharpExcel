using SharpExcel.Models.Styling.Borders;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.Models.Styling.Text;

namespace SharpExcel.Models.Styling;

public static class SharpExcelCellStyleConstants
{
    public static SharpExcelCellStyle DefaultHeaderStyle =
        new(SharpExcelColorConstants.Black, SharpExcelColorConstants.Transparent, TextStyle.Bold, borders: BorderCollection.HeaderDefault);
    
    public static SharpExcelCellStyle DefaultDataStyle =
        new(SharpExcelColorConstants.Black, SharpExcelColorConstants.Transparent, TextStyle.None, borders: BorderCollection.DataDefault);
}