using System.ComponentModel.DataAnnotations;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using SharpExcel.Abstraction;
using SharpExcel.Extensions;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Configuration;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Exporters;

/// <summary>
/// Base class for creating excel workbooks
/// </summary>
public class BaseExcelExporter<TModel> : IExcelExporter<TModel>
    where TModel : class, new()
{
    private readonly IOptions<ExporterOptions<TModel>> _options;

    public BaseExcelExporter(IOptions<ExporterOptions<TModel>> options)
    {
        _options = options;
    }
    /// <inheritdoc />
    public async Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null)
    {
        var parsedWorkbook = await ReadWorkbookAsync(sheetName, workbook, cultureInfo);
        ExporterHelpers.ApplyCellValidation(sheetName, workbook, parsedWorkbook);
        return workbook;
    }


    /// <inheritdoc />
    public async Task<XLWorkbook> GenerateWorkbookAsync(SharpExcelArguments arguments, IEnumerable<TModel> data, CultureInfo? cultureInfo = null)
    {
        var workbook = new XLWorkbook();
        
        var dropdownDataSheetName = ExcelParseHelper.GetDropdownDataSheetName();
        var worksheet = workbook.AddWorksheet(arguments.SheetName);

        //add extra hidden sheet where we can put data to show in enum dropdowns
        var dropdownWorksheet = workbook.AddWorksheet(dropdownDataSheetName).Hide();
        
        var headerStyle = _options.Value.Styling.DefaultHeaderStyle;

        if (headerStyle.RowHeight.HasValue)
        {
            worksheet.Rows().Height = headerStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var propertyMappings = TypeMapper.GetModelMetaData<TModel>();
        var dropdownDataMappings = EnumExporter.AddEnumDropdownMappingsToSheet(propertyMappings, dropdownWorksheet);
        var stylingRuleLookup = _options.Value.Styling.ToStylingRuleLookup();

        int offsetColumns = 0;
        WriteHeaderRow(propertyMappings, offsetColumns, worksheet, rowIndex, headerStyle);

        //go to next row to start inserting data
        rowIndex++;
        
        foreach (var dataItem in data)
        {
            var dataOffset = 0;
            WriteDataRow(propertyMappings, worksheet, rowIndex, dataOffset, stylingRuleLookup, dataItem, dropdownDataMappings, dropdownWorksheet);
            rowIndex++;
        }
        
        return workbook;
    }

    private void WriteDataRow(PropertyDataCollection propertyMappings, IXLWorksheet worksheet, int rowIndex, 
        int dataOffset, 
        Dictionary<string, List<StylingRule<TModel>>> stylingRuleLookup, TModel dataItem,
        Dictionary<Type, string> dropdownDataMappings, 
        IXLWorksheet dropdownWorksheet)
    {
        for (var i = 0; i < propertyMappings.PropertyMappings.Count; i++)
        {
            var mapping = propertyMappings.PropertyMappings[i];
            var row = worksheet.Row(rowIndex);
            var cell = worksheet.Cell(rowIndex, i + 1 - dataOffset /* use +1 because Excel starts at 1 */);
                
            var dataStyle = _options.Value.Styling.DefaultDataStyle;

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
    }

    private static void WriteHeaderRow(PropertyDataCollection propertyMappings,  int offsetColumns,
        IXLWorksheet worksheet, int rowIndex, SharpExcelCellStyle headerStyle)
    {
        for (var columnIndex = 0; columnIndex < propertyMappings.PropertyMappings.Count; columnIndex++)
        {
            var mapping = propertyMappings.PropertyMappings[columnIndex];
            
            var cell = worksheet.Cell(rowIndex, columnIndex + 1 - offsetColumns /* use +1 because Excel starts at 1 */);
            cell.Style.ApplyStyle(headerStyle);

            if (mapping.ColumnWidth > 0)
            {
                worksheet.Column(columnIndex + 1).Width = mapping.ColumnWidth;
            }

            cell.SetValue(mapping.Name);
        }
    }


    /// <inheritdoc />
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook, CultureInfo? cultureInfo = null)
    {
        var output = new ExcelReadResult<TModel>();
        var propertyData = TypeMapper.GetModelMetaData<TModel>();

        var sheet = workbook.Worksheet(sheetName);
        var usedArea = sheet.RangeUsed();

        //find header names based on TModel
        var headerNames = new HashSet<string>(propertyData.PropertyMappings.Where(
                x => !string.IsNullOrWhiteSpace(x.NormalizedName))
                .Select(x => x.NormalizedName?.ToLowerInvariant())!
            );

        //find header row
        var headerRowIndex = usedArea
            .Rows(x => x.Cells()
                .Any(c => headerNames.Contains(c.Value.ToString().ToLowerInvariant())))
            .FirstOrDefault()
            ?.RowNumber() ?? -1;

        var propertiesByColumnName = propertyData.PropertyMappings.ToDictionary(x => x.NormalizedName);

        foreach (var cell in sheet.Row(headerRowIndex).Cells())
        {
            if (!cell.TryGetValue(out string cellValue))
                continue;

            cellValue = cellValue.Trim().ToLowerInvariant();
            
            if (!headerNames.Contains(cellValue)) 
                continue;
            
            if (propertiesByColumnName.ContainsKey(cellValue))
            {
                propertiesByColumnName[cellValue].ColumnData = new()
                {
                    ColumnName = propertiesByColumnName[cellValue].Name ?? string.Empty,
                    ColumnIndex = cell.Address.ColumnNumber
                };
            }
        }
        
        var remainingRows = usedArea.Rows(headerRowIndex + 1, usedArea.RowCount()).ToList();

        //parse remaining data rows
        foreach (var row in remainingRows)
        {
            var data = new TModel();

            Dictionary<ExcelAddress, List<ValidationResult>> validationResults = new();

            foreach (var columnData in  propertyData.PropertyMappings.OrderBy(x => x.ColumnData.ColumnIndex))
            {
                var cell = row.Cell(columnData.ColumnData.ColumnIndex);

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