namespace SharpExcel.Models.Styling.Borders;

public class BorderCollection
{
    public BorderStyle[] Borders { get; set; } = new BorderStyle[4];

    public static BorderCollection DataDefault => new()
    {
        Borders =
        [
            BorderStyle.None,
            BorderStyle.None,
            BorderStyle.None,
            BorderStyle.None,
        ]
    };
    
    public static BorderCollection HeaderDefault => new()
    {
        Borders =
        [
            BorderStyle.None,
            BorderStyle.None,
            BorderStyle.Thick,
            BorderStyle.None
        ]
    };

    public BorderStyle GetBorderStyle(BorderDirection direction)
    {
        return Borders[(int)direction];
    }
    
    public BorderStyle SetBorderStyle(BorderDirection direction, BorderStyle data)
    {
        return Borders[(int)direction] = data;
    }
}