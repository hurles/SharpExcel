using ExcelSharp.Styling.Borders;
using ExcelSharp.Styling.Colorization;
using ExcelSharp.Styling.Text;

namespace ExcelSharp.Styling;

/// <summary>
/// Struct defining the styling of a cell
/// </summary>
public struct ExcelSharpCellStyle
{
    public ExcelSharpCellStyle(ExcelSharpColor textColor, ExcelSharpColor backgroundColor, TextStyle textStyle, float? rowHeight = null, float? fontSize = null,
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
    public ExcelSharpColor? TextColor { get; set; } = ExcelSharpColorConstants.Black;
    
    /// <summary>
    /// Background color used for cell
    /// if null, default color is used
    /// </summary>
    public ExcelSharpColor? BackgroundColor { get; set; } = ExcelSharpColorConstants.Black;

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