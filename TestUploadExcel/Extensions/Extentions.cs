using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using TestUploadExcel.Models;

namespace TestUploadExcel.Extensions
{
    public class ExcelMap
    {
        public string Name { get; set; }
        public string MappedTo { get; set; }
        public int Index { get; set; }
    }
    public class Error
    {
        public string PropertyName { get; set; }
        public string ErrorDescription { get; set; }
        public long RowNumber { get; set; }
    }
    public static class Extentionss
    {
        public static string[] GetHeaderColumns(this ExcelWorksheet sheet)
        {
            List<string> columnNames = new List<string>();
            foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
                columnNames.Add(firstRowCell.Text);
            return columnNames.ToArray();
        }
        //public static IEnumerable<T> ToList<T>(this ExcelWorksheet worksheet, out List<Error> errors) where T : new()
        //{
        //    errors = null;
        //    var propsInfo = TypeDescriptor.GetProperties(typeof(T)).Cast<PropertyDescriptor>().ToList();
        //    Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(Column);
        //    string[] headerValues = worksheet.GetHeaderColumns();

        //    using (var headers = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
        //    {
        //        if (headers.Columns != propsInfo.Count)
        //        {

        //        }
        //        if (!propsInfo.All(e => headers.Any(x => x.Value.Equals(e))))
        //        {

        //        }
        //        for (int i = 0; i < propsInfo.Count && i < headers.Count(); i++)
        //        {
        //            if (headerValues[i] == propsInfo[i].DisplayName)
        //            {
        //                System.Diagnostics.Debug.WriteLine($"Excel Header = ${headerValues[i]}  ---- Property Display Name = ${propsInfo[i].DisplayName}");
        //            }
        //        }

        //    }

        //    var columns = typeof(T)
        //            .GetProperties()
        //            .Where(x => x.CustomAttributes.Any(columnOnly))
        //    .Select(p => new
        //    {
        //        Property = p,
        //        Column = p.GetCustomAttributes<Column>().First().ColumnIndex //safe because if where above
        //    }).ToList();


        //    var rows = worksheet.Cells
        //        .Select(cell => cell.Start.Row)
        //        .Distinct()
        //        .OrderBy(x => x);


        //    //Create the collection container
        //    var collection = rows.Skip(1)
        //        .Select(row =>
        //        {
        //            var tnew = new T();
        //            columns.ForEach(col =>
        //            {
        //                //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
        //                //var value = worksheet.Cells[row, col.Column].Value.ToString();
        //                var val = worksheet.Cells[row, col.Column];
        //                //If it is numeric it is a double since that is how excel stores all numbers
        //                //if (val.Value == null)
        //                //{
        //                //    col.Property.SetValue(tnew, null);
        //                //    return;
        //                //}
        //                //if (col.Property.PropertyType == typeof(Int32))
        //                //{
        //                //    col.Property.SetValue(tnew, val.GetValue<int>());
        //                //    return;
        //                //}
        //                //if (col.Property.PropertyType == typeof(double))
        //                //{
        //                //    col.Property.SetValue(tnew, val.GetValue<double>());
        //                //    return;
        //                //}
        //                //if (col.Property.PropertyType == typeof(DateTime))
        //                //{
        //                //    col.Property.SetValue(tnew, val.GetValue<DateTime>());
        //                //    return;
        //                //}
        //                //Its a string
        //                col.Property.SetValue(tnew, val.GetValue<string>());

        //            });

        //            return tnew;
        //        });


        //    //Send it back
        //    return collection;
        //}

        public static List<T> ToList<T>(this ExcelWorksheet worksheet, out List<Error> errors) where T : new()
        {
            errors = new List<Error>();
            var propsInfo = TypeDescriptor.GetProperties(typeof(T)).Cast<PropertyDescriptor>().ToList();
            Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(Column);
            string[] headerValues = worksheet.GetHeaderColumns();

            using (var headers = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                if (headers.Columns != propsInfo.Count)
                {

                }
                for (int i = 0; i < propsInfo.Count && i < headers.Count(); i++)
                {
                    if (headerValues[i] == propsInfo[i].DisplayName)
                    {
                        System.Diagnostics.Debug.WriteLine($"Excel Header = ${headerValues[i]}  ---- Property Display Name = ${propsInfo[i].DisplayName}");
                    }
                }

            }

            var columns = typeof(T)
                    .GetProperties()
                    .Where(x => x.CustomAttributes.Any(columnOnly))
            .Select(p => new
            {
                Property = p,
                Column = p.GetCustomAttributes<Column>().First().ColumnIndex ,
                Attr = TypeDescriptor.GetProperties(typeof(T)).Cast<PropertyDescriptor>().ToList()
        }).ToList();

            var retList = new List<T>();
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;
            var startCol = start.Column;
            var startRow = start.Row;
            var endCol = end.Column;
            var endRow = end.Row;

            for (int rowIndex = startRow + 1; rowIndex <= endRow; rowIndex++)
            {
                var tnew = new T();
                columns.ForEach(col =>
                {
                    var val = worksheet.Cells[rowIndex, col.Column];
                    //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                    //                //var value = worksheet.Cells[row, col.Column].Value.ToString();
                    //                var val = worksheet.Cells[row, col.Column];
                    //                //If it is numeric it is a double since that is how excel stores all numbers
                    if (val.Value == null)
                    {
                        col.Property.SetValue(tnew, null);
                        return;
                    }
                    //                //if (col.Property.PropertyType == typeof(Int32))
                    //                //{
                    //                //    col.Property.SetValue(tnew, val.GetValue<int>());
                    //                //    return;
                    //                //}
                    //                //if (col.Property.PropertyType == typeof(double))
                    //                //{
                    //                //    col.Property.SetValue(tnew, val.GetValue<double>());
                    //                //    return;
                    //                //}
                    //                //if (col.Property.PropertyType == typeof(DateTime))
                    //                //{
                    //                //    col.Property.SetValue(tnew, val.GetValue<DateTime>());
                    //                //    return;
                    //                //}
                    //                //Its a string
                    col.Property.SetValue(tnew, val.GetValue<string>());
                });
                var context = new ValidationContext(tnew, serviceProvider: null, items: null);
                var results = new List<ValidationResult>();
                var isValid = Validator.TryValidateObject(tnew, context, results);
                if (!isValid)
                {
                    foreach (var validationResult in results)
                    {
                        errors.Add(new Error { PropertyName = validationResult.MemberNames.First().ToString(), ErrorDescription = validationResult.ErrorMessage ,RowNumber =rowIndex});
                    }
                }
                retList.Add(tnew);
            }
            
            return retList;
        }



    }
}
