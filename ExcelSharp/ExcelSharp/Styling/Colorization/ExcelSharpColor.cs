namespace ExcelSharp.Styling.Colorization;
public readonly struct ExcelSharpColor : IEquatable<ExcelSharpColor>
{
    public ExcelSharpColor(byte r, byte g, byte b, byte a = 255)
    {
        R = r;
        G = g;
        B = b;
        A = a;
    }

    private readonly byte[] _colorBytes = [0,0,0,255];
    
    public byte R
    {
        get => _colorBytes[0];
        set => _colorBytes[0] = value;
    }

    public byte G
    {
        get => _colorBytes[1];
        set => _colorBytes[1] = value;
    }

    public byte B
    {
        get => _colorBytes[2];
        set => _colorBytes[2] = value;
    }

    public byte A
    {
        get => _colorBytes[3];
        set => _colorBytes[3] = value;
    }

    public ExcelSharpColor WithAlpha(byte alpha)
    {
        return new ExcelSharpColor(_colorBytes[0], _colorBytes[1], _colorBytes[2], alpha);
    }
    
    public bool Equals(ExcelSharpColor other)
    {
        return _colorBytes.SequenceEqual(other._colorBytes);
    }
}