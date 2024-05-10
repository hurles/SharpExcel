using SharpExcel.Models.Styling.Borders;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.Models.Styling.Text;

namespace SharpExcel.Models.Styling;

/// <summary>
/// Struct defining the styling of a cell
/// </summary>
public struct SharpExcelCellStyle
{
    public SharpExcelCellStyle(SharpExcelColor textColor, SharpExcelColor backgroundColor, TextStyle textStyle, float? rowHeight = null, float? fontSize = null,
        BorderCollection? borders = null)
    {
        Borders = borders ?? BorderCollection.DataDefault;
        TextColor = textColor;
        BackgroundColor = backgroundColor;
        TextStyle = textStyle;
        RowHeight = rowHeight;
        FontSize = fontSize;
    }

    /// <summary>
    /// Color used for text within a cell
    /// if null, default color is used
    /// </summary>
    public SharpExcelColor? TextColor { get; set; } = SharpExcelColorConstants.Black;
    
    /// <summary>
    /// Background color used for cell
    /// if null, default color is used
    /// </summary>
    public SharpExcelColor? BackgroundColor { get; set; } = SharpExcelColorConstants.Black;

    /// <summary>
    /// Text style (Bold, Italic, etc)
    /// Default: TextStyle.None
    /// </summary>
    public TextStyle TextStyle { get; set; } = TextStyle.None;

    /// <summary>
    /// Height of cell
    /// if null, default height is used
    /// </summary>
    public double? RowHeight { get; set; }

    /// <summary>
    /// font size
    /// if null, default size is used
    /// </summary>
    public double? FontSize { get; set; } = null;

    /// <summary>
    /// Border styling
    /// </summary>
    public BorderCollection Borders { get; set; } = BorderCollection.DataDefault;
}