namespace SharpExcel.Extensions;

internal static class ExcelParseHelper
{
    public static string GetDropdownDataSheetName()
    {
        return GetName();
    }

    static string GetName()
    {
        var random = new Random();
        var randomNumber = random.Next(0, 1000000);
        return "Dropdowns_" + randomNumber.ToString("000000");
    }
}