using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;
using SharpExcel.Extensions;
using SharpExcel.Abstraction;
using SharpExcel.Models;
using SharpExcel.Models.Attributes;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling;

namespace SharpExcel;

/// <summary>
/// Base class for creating excel workbooks
/// </summary>
public abstract class BaseExcelExporter<TModel> : IExcelExporter<TModel>
    where TModel : class, new()
{
    /// <summary>
    /// method to actually build workbook
    /// </summary>
    /// <param name="arguments"></param>
    /// <param name="optionalColumnFunc">Functions that returns a boolean based on Property name of TModel (not column name), indicating whether or not to write the specified property based on a condition</param>
    /// <returns></returns>
    public async Task<XLWorkbook> BuildWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>>? optionalColumnFunc = null)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(arguments.SheetName);

        var headerStyle = OnSetHeaderStyle();

        if (headerStyle.RowHeight.HasValue)
        {
            worksheet.Rows().Height = headerStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var propertyMappings = GetModelMetaData();
        var optionalColumns = await GetOptionalPropertiesToExport(optionalColumnFunc);

        int offsetColumns = 0;
        for (var columnIndex = 0; columnIndex < propertyMappings.Count; columnIndex++)
        {

            var mapping = propertyMappings[columnIndex];

            if (mapping.Conditional && !optionalColumns.Contains(mapping.PropertyInfo.Name))
            {
                offsetColumns++;
                continue;
            }

            var cell = worksheet.Cell(rowIndex, columnIndex + 1 - offsetColumns /* use +1 because Excel starts at 1 */);
            
            
            cell.Style.ApplyStyle(headerStyle);

            if (mapping.ColumnWidth > 0)
            {
                worksheet.Column(columnIndex + 1).Width = mapping.ColumnWidth;
            }

            cell.SetValue(mapping.Name);
        }

        //go to next row to start inserting data
        rowIndex++;

        if (arguments?.Data is null)
            return workbook;

        foreach (var dataItem in arguments?.Data!)
        {
            var dataOffset = 0;
            for (var i = 0; i < propertyMappings.Count; i++)
            {
                var mapping = propertyMappings[i];
                var row = worksheet.Row(rowIndex);
                
                if (mapping.Conditional && !optionalColumns.Contains(mapping.PropertyInfo.Name))
                {
                    dataOffset++;
                    continue;
                }

                var cell = worksheet.Cell(rowIndex, i + 1 - dataOffset /* use +1 because Excel starts at 1 */);

                var dataStyle = OnSetCellDataStyle(mapping.PropertyInfo.Name, dataItem);

                if (dataStyle.RowHeight.HasValue &&  row.Height < dataStyle.RowHeight)
                {
                    row.Height = dataStyle.RowHeight.Value;
                }

                var dataValue = mapping.PropertyInfo.GetValue(dataItem);
                if (mapping.Format != null)
                {
                    if (dataValue is IFormattable formattable)
                    {
                        cell.SetValue(formattable.ToString(mapping.Format, CultureInfo.InvariantCulture));
                    }
                }
                cell.SetValue(XLCellValue.FromObject(dataValue));
                
                cell.Style.ApplyStyle(dataStyle);
            }

            rowIndex++;
        }
        
        return workbook;
    }

    public async Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(string sheetName, XLWorkbook workbook)
    {
        var parsedWorkbook = await ReadWorkbookAsync(sheetName, workbook);

        foreach (var result in parsedWorkbook.ValidationResults)
        {
            var cell = workbook.Worksheet(sheetName).Cell(result.Value.Address.RowNumber, result.Value.Address.ColumnId);
            var stringBuilder = new StringBuilder();
            foreach (var item in result.Value.ValidationResults)
            {
                stringBuilder.AppendLine(item.ErrorMessage);
            }
            cell.Style.ApplyStyle(SharpExcelCellStyleConstants.DefaultErrorStyle);
            cell.CreateComment().AddText(stringBuilder.ToString());
        }

        return workbook;
    }

    /// <summary>
    /// Reads a workbook to convert it into the given model
    /// </summary>
    /// <param name="sheetName">name of the sheet to read from</param>
    /// <param name="workbook"></param>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook)
    {
        var output = new ExcelReadResult<TModel>();
        var propertyData = GetModelMetaData();

        var sheet = workbook.Worksheet(sheetName);
        var usedArea = sheet.RangeUsed();

        //find header names based on TModel
        var headerNames = new HashSet<string>(propertyData.Where(x => !string.IsNullOrWhiteSpace(x.Name)).Select(x => x.Name)!);

        //find header row (so we can skip comments, etc)
        var headerRowIndex = usedArea
            .Rows(x => x.Cells()
                .All(c => headerNames.Contains(c.Value.ToString())))
            .FirstOrDefault()
            ?.RowNumber() ?? -1;

        var remainingRows = usedArea.Rows(headerRowIndex + 1, usedArea.RowCount()).ToList();

        //parse remaining data rows
        foreach (var row in remainingRows)
        {
            var data = new TModel();

            Dictionary<ExcelAddress, List<ValidationResult>> validationResults = new();

            for (var columnIndex = 0; columnIndex < propertyData.Count; columnIndex++)
            {
                var columnData = propertyData[columnIndex];
                var cell = row.Cell(columnIndex + 1 /* use +1 because Excel starts at 1 */);

                var excelAddress = new ExcelAddress()
                {
                    RowNumber = row.RowNumber(),
                    ColumnId = cell.Address.ColumnNumber,
                    ColumnName = cell.Address.ColumnLetter,
                    HeaderName = columnData.Name
                };
                
                var dataValue = 
                    //fp types
                    TrySetValue<double>(columnData, cell) ??
                    TrySetValue<float>(columnData, cell) ??
                    TrySetValue<decimal>(columnData, cell) ??
                    
                    //integer types
                    TrySetValue<int>(columnData, cell) ??
                    TrySetValue<uint>(columnData, cell) ??
                    TrySetValue<byte>(columnData, cell) ??
                    TrySetValue<short>(columnData, cell) ??
                    TrySetValue<ushort>(columnData, cell) ??
                    
                    TrySetValue<long>(columnData, cell) ??
                    TrySetValue<ulong>(columnData, cell) ??
                    
                    //dates
                    TrySetValue<DateTime>(columnData, cell) ??
                    TrySetValue<DateTimeOffset>(columnData, cell) ??
                    
                    TrySetValue<string>(columnData, cell);

                if (columnData.PropertyInfo.PropertyType == dataValue?.GetType())
                {
                    columnData.PropertyInfo.SetValue(data, dataValue);
                    
                    ValidationContext context = new ValidationContext(data) { 
                        MemberName = columnData.PropertyInfo.Name, 
                        DisplayName = columnData.Name ?? columnData.PropertyInfo.Name};
                    var validations = new List<ValidationResult>();
                    if (!Validator.TryValidateProperty(dataValue, context, validations))
                    {
                        validationResults.Add(excelAddress, validations);
                    }
                }
            }
            
            output.Records.Add(data);
            if (validationResults.Any())
            {
                foreach (var validationResult in validationResults)
                {
                    output.ValidationResults.Add(data, new ExcelCellValidationResult()
                    {
                        Address = validationResult.Key,
                        ValidationResults = validationResult.Value
                    });
                }
            }
            
        }

        return Task.FromResult(output);
    }

    /// <summary>
    /// Override this method to set cell style for header row cells.
    /// </summary>
    /// <returns></returns>
    protected virtual SharpExcelCellStyle OnSetHeaderStyle()
    {
        return SharpExcelCellStyleConstants.DefaultHeaderStyle;
    }

    /// <summary>
    /// Override this method to set cell style for each data cell.
    /// Record + current column are provided so styling can be different based on conditions given by the user
    /// </summary>
    /// <param name="record">current record being processed</param>
    /// <param name="propertyName">current column being processed</param>
    /// <returns></returns>
    protected virtual SharpExcelCellStyle OnSetCellDataStyle(string propertyName, TModel record)
    {
        return SharpExcelCellStyleConstants.DefaultDataStyle;
    }

    public async Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>>? conditionalColumnFunc = null)
    {
        var output = new List<string>();

        var conditionalProperties = GetModelMetaData().Where(x => x.Conditional);

        foreach (var optional in conditionalProperties)
        {
            if (conditionalColumnFunc != null && await conditionalColumnFunc(optional.PropertyInfo.Name))
            {
                output.Add(optional.PropertyInfo.Name);
            }
        }
        
        return [..output];
    }

    public virtual HashSet<string> GetOptionalColumns()
    {
        var propertyData = GetModelMetaData();
        var output = propertyData.Where(x => x.Conditional)
            .Select(x => x.Name);

        return [..output];
    }

    static object? TrySetValue<TPropertyData>(PropertyData columnData, IXLCell cell)
    {
        if (columnData.PropertyInfo.PropertyType != typeof(TPropertyData)) 
            return null;
        
        if (cell.TryGetValue(out TPropertyData dataValue))
        {
            return dataValue;
        }

        return null;
    }

    /// <summary>
    /// Reads model attributes and converts to column metadata
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    private static List<PropertyData> GetModelMetaData()
    {
        var dataType = typeof(TModel);
        var propertyMappings = new List<PropertyData>();
        for (var columnIndex = 0; columnIndex < dataType.GetProperties().Length; columnIndex++)
        {

            var property = dataType.GetProperties()[columnIndex];

            var attribute = property.GetCustomAttribute<ExcelColumnDefinitionAttribute>();

            if (attribute is null)
            {
                continue;
            }

            var columnName = property.Name;
            if (!string.IsNullOrWhiteSpace(attribute?.DisplayName))
            {
                columnName = attribute?.DisplayName;
            }

            propertyMappings.Add(new PropertyData()
            {
                Name = columnName,
                PropertyInfo = property,
                Format = attribute?.Format,
                Conditional = attribute?.IsConditional ?? false,
                ColumnWidth = attribute?.ColumnWidth ?? -1
            });
        }

        return propertyMappings;
    }
    
    /// <summary>
    /// This struct is only used to load the metadata of the model
    /// </summary>
    private struct PropertyData
    {
        public string? Name { get; set; }

        public string? Format { get; set; }

        public int ColumnWidth { get; set; }

        public bool Conditional { get; set; }

        public PropertyInfo PropertyInfo { get; set; }
    }
}