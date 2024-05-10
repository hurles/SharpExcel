using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using SharpExcel.Abstraction;
using SharpExcel.Extensions;
using SharpExcel.Models;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling;

namespace SharpExcel.Exporters;

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
    /// <param name="cultureInfo">Culture used to generate workbook</param>
    /// <returns></returns>
    public async Task<XLWorkbook> GenerateWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>>? optionalColumnFunc = null, CultureInfo? cultureInfo = null)
    {
        var workbook = new XLWorkbook();
        
        var dropdownDataSheetName = ExcelParseHelper.GetDropdownDataSheetName();
        var worksheet = workbook.AddWorksheet(arguments.SheetName);

        //add extra hidden sheet where we can put data to show in enum dropdowns
        var dropdownWorksheet = workbook.AddWorksheet(dropdownDataSheetName);
        //dropdownWorksheet.Hide();
        
        var headerStyle = arguments.StylingCollection.DefaultHeaderStyle;

        if (headerStyle.RowHeight.HasValue)
        {
            worksheet.Rows().Height = headerStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var propertyMappings = TypeMapper.GetModelMetaData<TModel>();
        
        var optionalColumns = await GetOptionalPropertiesToExport(optionalColumnFunc);
        
        var dropdownDataMappings = EnumExporter.AddEnumDropdownMappings(propertyMappings, dropdownWorksheet);

        var stylingRuleLookup = arguments.StylingCollection.ToStylingRuleLookup();

        int offsetColumns = 0;
        for (var columnIndex = 0; columnIndex < propertyMappings.PropertyMappings.Count; columnIndex++)
        {
            var mapping = propertyMappings.PropertyMappings[columnIndex];

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
            for (var i = 0; i < propertyMappings.PropertyMappings.Count; i++)
            {
                var mapping = propertyMappings.PropertyMappings[i];
                var row = worksheet.Row(rowIndex);
                
                if (mapping.Conditional && !optionalColumns.Contains(mapping.PropertyInfo.Name))
                {
                    dataOffset++;
                    continue;
                }

                var cell = worksheet.Cell(rowIndex, i + 1 - dataOffset /* use +1 because Excel starts at 1 */);
                
                var dataStyle = arguments.StylingCollection.DefaultDataStyle;

                if (stylingRuleLookup.TryGetValue(mapping.PropertyInfo.Name, out var rules))
                {
                    foreach (var rule in rules)
                    {
                        var ruleStyle = rule.EvaluateRules(dataItem);
                        dataStyle = ruleStyle ?? dataStyle;
                    }
                }

                if (dataStyle.RowHeight.HasValue && row.Height < dataStyle.RowHeight)
                {
                    row.Height = dataStyle.RowHeight.Value;
                }

                var dataValue = mapping.PropertyInfo.GetValue(dataItem);
                
                //handle enums
                if (mapping.PropertyInfo.PropertyType.IsEnum)
                {
                    EnumExporter.WriteEnumValue(propertyMappings, mapping, dataValue, cell, dropdownDataMappings, dropdownWorksheet);
                }
                //handle format
                else if (mapping.Format != null)
                {
                    if (dataValue is IFormattable formattable)
                    {
                        cell.SetValue(formattable.ToString(mapping.Format, CultureInfo.InvariantCulture));
                    }
                }
                else
                {
                    cell.SetValue(XLCellValue.FromObject(dataValue));
                }
                
                cell.Style.ApplyStyle(dataStyle);
            }

            rowIndex++;
        }
        
        return workbook;
    }

    public async Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null)
    {
        var parsedWorkbook = await ReadWorkbookAsync(sheetName, workbook, cultureInfo);

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
    /// <param name="cultureInfo">culture used, defaults to CurrentCulture if null</param>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null)
    {
        var output = new ExcelReadResult<TModel>();
        var propertyData = TypeMapper.GetModelMetaData<TModel>();

        var sheet = workbook.Worksheet(sheetName);
        var usedArea = sheet.RangeUsed();

        //find header names based on TModel
        var headerNames = new HashSet<string>(propertyData.PropertyMappings.Where(x => !string.IsNullOrWhiteSpace(x.Name)).Select(x => x.Name)!);

        //find header row
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

            for (var columnIndex = 0; columnIndex < propertyData.PropertyMappings.Count; columnIndex++)
            {
                var columnData = propertyData.PropertyMappings[columnIndex];
                var cell = row.Cell(columnIndex + 1 /* use +1 because Excel starts at 1 */);

                var excelAddress = new ExcelAddress()
                {
                    RowNumber = row.RowNumber(),
                    ColumnId = cell.Address.ColumnNumber,
                    ColumnName = cell.Address.ColumnLetter,
                    HeaderName = columnData.Name
                };
                
                var dataValue = TrySetValue(propertyData, columnData, cell, cultureInfo ?? CultureInfo.CurrentCulture);

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

    public async Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>>? conditionalColumnFunc = null)
    {
        var output = new List<string>();

        var conditionalProperties = TypeMapper.GetModelMetaData<TModel>().PropertyMappings.Where(x => x.Conditional);

        foreach (var optional in conditionalProperties)
        {
            if (conditionalColumnFunc != null && await conditionalColumnFunc(optional.PropertyInfo.Name))
            {
                output.Add(optional.PropertyInfo.Name);
            }
        }
        
        return [..output];
    }

    static object? TrySetValue(PropertyDataCollection dataCollection, PropertyData columnData, IXLCell cell, CultureInfo cultureInfo)
    {
        //extract underlying nullable type if there is one
        var actualType = Nullable.GetUnderlyingType(columnData.PropertyInfo.PropertyType) ?? columnData.PropertyInfo.PropertyType;
        
        //handle numeric types
        if (actualType.IsNumeric())
        {
            if (cell.TryGetValue(out decimal numericValue))
            {
                return Convert.ChangeType(numericValue, actualType);
            }
        }

        //handle booleans
        if (actualType == typeof(bool))
        {
            if (cell.TryGetValue(out bool booleanValue))
            {
                return booleanValue;
            }
        }

        //handle strings
        if (actualType == typeof(string))
        {
            if (cell.TryGetValue(out string textValue))
            {
                return textValue;
            }

            var displayVal = cell.CachedValue.ToString(cultureInfo);
            
            if (!string.IsNullOrWhiteSpace(displayVal))
            {
                return displayVal;
            }
        }

        if (actualType.IsEnum)
        {
            if (cell.TryGetValue(out string textValue))
            {
                if (dataCollection.EnumMappings.TryGetValue(actualType, out var data))
                {
                    var value = data.FirstOrDefault(x => x.VisualName?.ToLowerInvariant() == textValue.Trim().ToLowerInvariant());
                    return value.Value;
                }
            }
            
        }
        
        return default;
    }
}