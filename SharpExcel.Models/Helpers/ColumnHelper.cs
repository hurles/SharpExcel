namespace SharpExcel.Models.Helpers;

public static class ColumnHelper
{
    public static string ConvertToAddress(int column, int row)
    {
        string columnAddress = ConvertToColumnAddress(column);
        return $"{columnAddress}{row}";
    }

    private static string ConvertToColumnAddress(int column)
    {
        string columnAddress = string.Empty;
        while (column > 0)
        {
            column--;
            columnAddress = (char)('A' + (column % 26)) + columnAddress;
            column /= 26;
        }

        return columnAddress;
    }
    
    public static (int column, int row) ConvertFromAddress(string address)
    {
        int rowIndex = 0;
        int columnIndex = 0;

        // Find the position where the row number starts
        int i = 0;
        while (i < address.Length && !char.IsDigit(address[i]))
        {
            i++;
        }

        // Extract the column letters and row number
        string columnLetters = address.Substring(0, i);
        string rowNumber = address.Substring(i);

        // Convert column letters to column index
        for (int j = 0; j < columnLetters.Length; j++)
        {
            columnIndex *= 26;
            columnIndex += (columnLetters[j] - 'A' + 1);
        }

        // Convert row number to row index
        rowIndex = int.Parse(rowNumber);

        return (columnIndex, rowIndex);
    }
}