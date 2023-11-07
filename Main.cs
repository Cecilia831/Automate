using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;

namespace Automate
{
    internal class Program
    {
        static void Main(string[] args)
        {

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            ExcelFile workbook = new ExcelFile();
            ExcelWorksheet worksheet = workbook.Worksheets.Add("Sheet1");
            ExcelCell cell = worksheet.Cells["A1"];

            cell.Value = "Hello World!";

            workbook.Save("HelloWorld.xlsx");
        } 
    }
}
