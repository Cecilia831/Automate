using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Windows.Documents;
using static System.Collections.Specialized.BitVector32;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static Automate.Program;

namespace Automate
{
    internal class Program
    {
        static void Main()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            BuildSheet();
            ReadInputRow();
            DisplaySheet();
            LogIn();
            var r = ReadInputRow();
            Console.WriteLine(r);
           
        }

        public static class Globals
        {
            public const Int32 N = 7; 
        }
       
        static void BuildSheet() {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();
            ExcelRow row = worksheet.Rows.First();
            ExcelCell cell = row.Cells.First();

            worksheet.Cells[0, 0].Value = "Title";
            worksheet.Cells[0 ,1].Value = "Assigned to";
            worksheet.Cells[0, 2].Value = "Title2";
            worksheet.Cells[0, 3].Value = "Cost Code";
            worksheet.Cells[0, 4].Value = "Unit Cost";
            worksheet.Cells[0, 5].Value = "Invoice Date";
            worksheet.Cells[0, 6].Value = "Due Date";
            workbook.Save("Input sheet.xlsx");
        }

        static IDictionary<string, string> ReadInputRow()
        {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();

            int i = 0;
            IDictionary<string, string> row = new Dictionary<string, string>();
            while (i < Globals.N)
            {
                row.Add(worksheet.Cells[0, i].ToString(), worksheet.Cells[1,i].ToString());
                i++;
            }
            return row;
        }

        static void DisplaySheet() {

            // Load Excel workbook from file's path.
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");

            // Iterate through all worksheets in a workbook.
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                // Display sheet's name.
                Console.WriteLine("{1} {0} {1}\n", worksheet.Name, new string('#', 30));

                // Iterate through all rows in a worksheet.
                foreach (ExcelRow row in worksheet.Rows)
                {
                    // Iterate through all allocated cells in a row.
                    foreach (ExcelCell cell in row.AllocatedCells)
                    {
                        // Read cell's data.
                        string value = cell.Value?.ToString() ?? "EMPTY";

                        // For merged cells, read only the first cell's data.
                        if (cell.MergedRange != null && cell.MergedRange[0] != cell)
                            value = "MERGED";

                        // Display cell's value and type.
                        value = value.Length > 15 ? value.Remove(15) + "…" : value;
                        Console.Write($"{value} [{cell.ValueType}]".PadRight(30));
                    }

                    Console.WriteLine();
                }
            }

            Console.WriteLine();
        }

        static void LogIn()
        {
            var d = new ChromeDriver();
            d.Navigate().GoToUrl("https://buildertrend.net/summaryGrid.aspx");
            //d.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            //user name:lisa@sprucebox.com
            //password:SB12345$
            d.FindElement(By.Id("username")).SendKeys("lisa@sprucebox.com");
            d.FindElement(By.Id("password")).SendKeys("SB12345$");
            var button = d.FindElement(By.ClassName("ant-btn-primary"));
            button.Click();
            //d.Quit();
        }
    }
}
