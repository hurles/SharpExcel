namespace SharpExcel.Models.Styling.Colorization;
public readonly struct SharpExcelColor : IEquatable<SharpExcelColor>
{
    public SharpExcelColor(byte r, byte g, byte b, byte a = 255)
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

    public SharpExcelColor WithAlpha(byte alpha)
    {
        return new SharpExcelColor(_colorBytes[0], _colorBytes[1], _colorBytes[2], alpha);
    }
    
    public bool Equals(SharpExcelColor other)
    {
        return _colorBytes.SequenceEqual(other._colorBytes);
    }
}