using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Helper
{
    public class DownloadXL
    {
        public static async Task<IActionResult> GetExcel<T>(string WorkSheetName, List<T> dt)
        {
            int totalRowCount = dt.Count;
            int totalColumnCount = 1;

            XLWorkbook xLWorkbook = new XLWorkbook();
            IXLWorksheet xLWorksheet = xLWorkbook.Worksheets.Add(WorkSheetName);

            //Title
            var Title = xLWorksheet.Range(2, 6, 2, 13).SetValue(WorkSheetName);
            Title.Merge().Style.Font.SetBold().Font.FontSize = 22;
            Title.Style.Font.SetFontColor(XLColor.CoolBlack);
            Title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            Title.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            string address = "Address: 6/2, Kazi Nazrul Islam Road, Mohammadpur";
            string Email = "Email: info@ibos.io";
            string Phone = "Phone: +123456678";
            string invoice = "Invoice";
            
            var subTitle = xLWorksheet.Range(3, 8, 3, 11).SetValue(address);
            subTitle.Merge().Style.Font.FontSize = 12;
            subTitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            subTitle.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var subTitle2 = xLWorksheet.Range(4, 8, 4, 11).SetValue(Email + "   " + Phone);
            subTitle2.Merge().Style.Font.FontSize = 12;
            subTitle2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            subTitle2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var subTitle3 = xLWorksheet.Range(6, 10, 6, 10).SetValue(invoice);
            subTitle3.Merge().Style.Font.SetBold().Font.SetFontSize(16).Font.SetFontColor(XLColor.CoolBlack);
            subTitle2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            subTitle3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            // Cell Border 
            subTitle3.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            subTitle3.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            subTitle3.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            subTitle3.Style.Border.LeftBorder = XLBorderStyleValues.Thin;


            // Outside Border of Title in a range
            //var border = xLWorksheet.Range(3, 8, 4, 11); 
            //border.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
           

            //Header (for one column, one row)
            var header = xLWorksheet.Range(8, 8, 8, 8);
            header.Style.Font.SetBold();
            header.Style.Fill.SetBackgroundColor(XLColor.CoolBlack);
            header.Style.Font.SetFontColor(XLColor.White);
            header.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            header.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            header.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            //header.Style.Border.BottomBorder(XLBorderStyleValues.Thin);
            header.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            header.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            header.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

            //int index = 1;   // setoutSideBorder in cell
            //for (int i = 0; i < totalColumnCount; i++)
            //{
            //    header.Cell(1, index).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            //    index++;
            //}

            // Column Name
            var hsl = 1;
            var TypeT = typeof(T); // GetExcel<T>(string WorkSheetName, List<T> dt) ==> GetExcel<ProductDTO>("abc", data) --> Generic method.
            var properties = TypeT.GetProperties(); // Get ColumnName 
            foreach (var item in properties)  
            {
                var name = StringSpliter.ByPascaleCase(item.Name);
                //Cell(row number, column number) of a Range
                header.Cell(1, hsl++).SetValue(name);  // Set ColumnName In cell [ 1 means 1st row in a Range of (7, 8, 7, 8)]
            }

            //Table Data
            
            var dataArray = xLWorksheet.Range(9, 8, totalRowCount,11 );  // range of cell for adding value into cell
            dataArray.Style.Font.SetBold(false);
            dataArray.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            var RowIndex = 1;
            foreach (var row in dt) // for row wise increment
            {
                var rsl = 1;
                foreach (var item in properties) // for Column wise increment
                {
                    var type = row.GetType();  // get type of property
                    var itemName = type.GetProperty(item.Name); // get columnName
                    var value = itemName.GetValue(row) ?? null; // get value of column
                    dataArray.Cell(RowIndex, rsl++).SetValue(value is null ? null : value.ToString());

                }

                RowIndex++;
            }
            xLWorksheet.Columns().AdjustToContents(); //for adjustment of column

            MemoryStream ms = new MemoryStream();
            xLWorkbook.SaveAs(ms);
            ms.Position = 0;

            return new FileStreamResult(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { FileDownloadName = $"{WorkSheetName} - {System.DateTime.Now}.xlsx" };
        }
    }
}
