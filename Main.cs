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

namespace Automate
{
    internal class Program
    {
        static void Main()
        {
            OpenSheet();
            LogIn();
           
        }

        static void LogIn() {
            var d = new ChromeDriver();
            d.Navigate().GoToUrl("https://buildertrend.net/summaryGrid.aspx");
            //user name:lisa@sprucebox.com
            //password:SB12345$
            d.FindElement(By.Id("username")).SendKeys("lisa@sprucebox.com");
            d.FindElement(By.Id("password")).SendKeys("SB12345$");
            var button = d.FindElement(By.ClassName("ant-btn-primary"));
            button.Click();
            d.Quit();
        }

        static void OpenSheet() {

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Load Excel workbook from file's path.
            ExcelFile workbook = ExcelFile.Load("Username Password.xlsx");

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
        }
    }
}
