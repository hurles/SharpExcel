using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
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
        
        var headerStyle = OnSetHeaderStyle();

        if (headerStyle.RowHeight.HasValue)
        {
            worksheet.Rows().Height = headerStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var propertyMappings = TypeMapper.GetModelMetaData<TModel>();
        var optionalColumns = await GetOptionalPropertiesToExport(optionalColumnFunc);
        
        var dropdownDataMappings = AddEnumDropdownMappings(propertyMappings, dropdownWorksheet);

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

                var dataStyle = OnSetCellDataStyle(mapping.PropertyInfo.Name, dataItem);

                if (dataStyle.RowHeight.HasValue &&  row.Height < dataStyle.RowHeight)
                {
                    row.Height = dataStyle.RowHeight.Value;
                }

                var dataValue = mapping.PropertyInfo.GetValue(dataItem);
                
                //handle enums
                if (mapping.PropertyInfo.PropertyType.IsEnum)
                {
                    WriteEnumValue(propertyMappings, mapping, dataValue, cell, dropdownDataMappings, dropdownWorksheet);
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

    private static void WriteEnumValue(PropertyDataCollection propertyMappings, PropertyData mapping, object dataValue,
        IXLCell cell, Dictionary<Type, string> dropdownDataMappings, IXLWorksheet dropdownWorksheet)
    {
        if (propertyMappings.EnumMappings.TryGetValue(mapping.PropertyInfo.PropertyType, out var enumValues))
        {
            var text = dataValue.ToString().Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(text))
            {
                foreach (var enumValue in enumValues)
                {
                    if (enumValue.Name == text)
                    {
                        cell.SetValue(enumValue.VisualName);
                    }
                }
            }
                        
            if (dropdownDataMappings.TryGetValue(mapping.PropertyInfo.PropertyType, out var range))
            {
                cell.CreateDataValidation().List(dropdownWorksheet.Range(range), true);
            }
        }
    }

    private static Dictionary<Type, string> AddEnumDropdownMappings(PropertyDataCollection propertyMappings, IXLWorksheet dropdownWorksheet)
    {
        int dropDownWorkbookColumn = 1;
        var dropdownDataMappings = new Dictionary<Type, string>();
        foreach (var enumMapping in propertyMappings.EnumMappings)
        {
            var columnLength = 0;
            for (int i = 0; i < enumMapping.Value.Count; i++)
            {
                var cell = dropdownWorksheet.Row(i + 1).Cell(dropDownWorkbookColumn);
                cell.SetValue(enumMapping.Value[i].VisualName);
                columnLength++;
            }
            var letter = dropdownWorksheet.Column(dropDownWorkbookColumn).ColumnLetter();
            dropdownDataMappings.Add(enumMapping.Key, $"{letter}{1}:{letter}{columnLength + 1}");
            dropDownWorkbookColumn++;
        }

        return dropdownDataMappings;
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

internal static class TypeMapper
{
        /// <summary>
    /// Reads model attributes and converts to column metadata
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public static PropertyDataCollection GetModelMetaData<TModel>() 
    {
        var propertyDataCollection = new PropertyDataCollection();
        
        var dataType = typeof(TModel);
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
            
            if (property.PropertyType.IsEnum)
            {
                if (!propertyDataCollection.EnumMappings.ContainsKey(property.PropertyType))
                {
                    propertyDataCollection.EnumMappings.Add(property.PropertyType, GetEnumMappings(property.PropertyType));
                }
            }

            propertyDataCollection.PropertyMappings.Add(new PropertyData()
            {
                Name = columnName,
                PropertyInfo = property,
                Format = attribute?.Format,
                Conditional = attribute?.IsConditional ?? false,
                ColumnWidth = attribute?.ColumnWidth ?? -1
            });
        }

        return propertyDataCollection;
    }

    private static List<EnumData> GetEnumMappings(Type propertyType)
    {
        if (!propertyType.IsEnum)
            return new();
        
        var enumDataList = new List<EnumData>();
        var memberInfos = propertyType.GetFields();

        foreach (var member in memberInfos.Where(m => m.FieldType.IsEnum))
        {
            var displayAttribute = member.GetCustomAttribute<DisplayAttribute>();
            object enumValue = Enum.ToObject(propertyType, member.GetRawConstantValue());
            
            var enumData = new EnumData()
            {
                VisualName = displayAttribute?.Name ?? member.Name,
                Name = member.Name.ToLowerInvariant(),
                //use long because technically you can use longs for enum values
                NumericValue = Convert.ToInt64(member.GetRawConstantValue()),
                Value = enumValue
            };
            enumDataList.Add(enumData);
        }

        return enumDataList;
    }
}