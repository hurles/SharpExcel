using SharpExcel.Models.Styling.Borders;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.Models.Styling.Text;

namespace SharpExcel.Models.Styling.Constants;

public static class ExcelCellStyleConstants
{
    public static ExcelCellStyle DefaultHeaderStyle =
        new(ExcelColorConstants.Black, ExcelColorConstants.Transparent, TextStyle.Bold, borders: BorderCollection.HeaderDefault);
    
    public static ExcelCellStyle DefaultDataStyle =
        new(ExcelColorConstants.Black, ExcelColorConstants.Transparent, TextStyle.None, borders: BorderCollection.DataDefault);
    
    public static ExcelCellStyle DefaultErrorStyle =
        new(new (80, 40, 40), new (255, 150, 150), TextStyle.None, borders: BorderCollection.DataDefault);
}