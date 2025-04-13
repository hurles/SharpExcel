using System.ComponentModel.DataAnnotations;

namespace SharpExcel.Models.Data;

public interface ITargetingRule
{
    
}
public record TargetingRule<TRecord> : ITargetingRule
{
    /// <summary>
    /// REQUIRED: name of the sheet in the excel file
    /// </summary>
    [MinLength(1)]
    public string SheetName { get; set; } = null!;

    /// <summary>
    /// Optional Row to start reading/writing from.
    /// This is useful when you want to only affect part of a sheet.
    /// Excel rows start as 1, so 1 is the first row
    /// </summary>
    public int? Row { get; set; }
    
    /// <summary>
    /// Optional Column to start reading/writing from.
    /// This is useful when you want to only affect part of a sheet.
    /// Excel columns start as 1, so 1 is the first row
    /// </summary>
    public int? Column { get; set; }
    
    /// <summary>
    /// Conditions to check if the rule should be applied.
    /// </summary>
    public Func<TRecord, bool>? RulePredicate { get; set; }
    
    public TargetingRule<TRecord> WithCondition(Func<TRecord, bool> condition)
    {
        RulePredicate = condition;
        //return this object so we can chain calls
        return this;
    }
    
    /// <summary>
    /// Sets the row to start reading/writing from.
    /// </summary>
    /// <param name="rowId"></param>
    /// <returns></returns>
    public TargetingRule<TRecord> WithStartRow(int rowId)
    {
        Row = rowId;
        //return this object so we can chain calls
        return this;
    }
    
    /// <summary>
    /// Sets the column to start reading/writing from.
    /// </summary>
    public TargetingRule<TRecord> WithStartColumn(int columnId)
    {
        Column = columnId;
        //return this object so we can chain calls
        return this;
    }
    
    /// <summary>
    /// Sets the sheet name to start reading/writing from.
    /// </summary>
    public TargetingRule<TRecord> WithSheetName(string sheetName)
    {
        SheetName = sheetName;
        //return this object so we can chain calls
        return this;
    }
}


