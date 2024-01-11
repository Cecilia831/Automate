using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Documents;
using GemBox.Spreadsheet;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

using static System.Collections.Specialized.BitVector32;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static Automate.Program;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using OpenQA.Selenium.Internal;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using System.Collections;

namespace Automate
{
    internal class Program
    {
        static void Main()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            //BuildSheet();
            //DisplaySheet();
            var Login = LogIn();
            var r = ReadInputRow();
            int ProNum = CheckProjectsNum();
            Console.WriteLine("{0} projects wait in line", ProNum - 1);
            while (ProNum > 1)
            {
                FinancialBillsPOs(Login);
                SearchAndNewPO(Login, r);
                Console.WriteLine("**********************************");
                foreach (KeyValuePair<string, string> ele in r)
                    Console.WriteLine("{0}: {1}", ele.Key, ele.Value);
                InputPO(Login, r);
                DeleteFromInputSheet();
                ClearSearchBoxGoBackSummary(Login);
                r = ReadInputRow();
                ProNum--;
            }
            Console.WriteLine("**Input sheet is empty. All Projects have entered!**");
        }

        public static class Globals
        {
            public const Int32 N = 8; 
        }
       
        static void BuildSheet() {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();
            ExcelRow row = worksheet.Rows.First();

            worksheet.Cells[0, 0].Value = "Project No";
            worksheet.Cells[0, 1].Value = "Title";
            worksheet.Cells[0 ,2].Value = "Assigned to";
            worksheet.Cells[0, 3].Value = "Title2";
            worksheet.Cells[0, 4].Value = "Cost Code";
            worksheet.Cells[0, 5].Value = "Unit Cost";
            worksheet.Cells[0, 6].Value = "Invoice Date";
            worksheet.Cells[0, 7].Value = "Due Date";
            workbook.Save("Input sheet.xlsx");
        }

        static IDictionary<string, string> ReadInputRow()
        {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();
            int i = 0;
            IDictionary<string, string> row = new Dictionary<string, string>();
            try
            {
                while (i < Globals.N)
                {
                    row.Add(worksheet.Cells[0, i].Value.ToString(), worksheet.Cells[1, i].Value.ToString());
                    i++;
                }
                return row;
            }
            catch {
                return null;
            }
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

        static ChromeDriver LogIn()
        {
            //Disabled all Chrome-level notifications
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--disable-extensions"); // to disable extension
            options.AddArguments("--disable-notifications"); // to disable notification
            options.AddArguments("--disable-application-cache"); // to disable cache
            options.AddArgument("--start-maximized"); // to maximize window
            options.DebuggerAddress = "127.0.0.1:9999";

            var d = new ChromeDriver(options);

            //d.Navigate().GoToUrl("https://buildertrend.net/summaryGrid.aspx");


            return d;
        }

        //From Summary Page, Goto Financial->Purchase Order
        static void FinancialBillsPOs(ChromeDriver d)
        {
            Thread.Sleep(5000);
            //Find Financial
            var b = d.FindElement(By.XPath("//html/body/div[2]/div/div/div[3]/form/div[3]/div[4]/div/div/div[1]/div/div[1]/div/div[6]/button"));
            b.Click();
            Thread.Sleep(2000);
            //Find Purchase Order
            var BP = d.FindElement(By.XPath("//*[@id=\"reactMainNavigation\"]/div/div[1]/div/div[6]/div/div/div/ul/li[3]/span/div/div/a/div/div/div[2]/div"));
            BP.Click();
            Thread.Sleep(5000);

            // Close the chatbox if possible
            try
            {
                //IFrame - Close ChatBox
                //Switch to the frame
                d.SwitchTo().Frame("intercom-launcher-frame");
                Thread.Sleep(3000);
                //Now click the button
                var e = d.FindElement(By.CssSelector("#intercom-container > div > div > div.intercom-1epm6qj.e11rlguj3 > svg"));
                e.Click();
                Thread.Sleep(1000);
                // Return to the top level
                d.SwitchTo().DefaultContent();
                Thread.Sleep(3000);
                //Click Close Button
                e = d.FindElement(By.CssSelector("#btnCloseIntercom"));
                e.Click();
                Thread.Sleep(3000);
            }
            catch
            {
            }
            finally
            {
                d.SwitchTo().DefaultContent();
            }
        }

        static int CheckProjectsNum() {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();
            int rows = worksheet.Rows.Count();
            return rows;
        }

        static void SearchAndNewPO(ChromeDriver d, IDictionary<String, String> row) {
            Thread.Sleep(2000);
            d.FindElement(By.Id("JobSearch")).SendKeys(row["Project No"]);
            Thread.Sleep(2000);
            d.FindElement(By.ClassName("ItemRowJobName")).Click();// Click to Job Order
            Thread.Sleep(5000);
            //Find and click New -> PO
            d.FindElement(By.CssSelector("#rc-tabs-0-panel-1 > div > div.GridContainer-Header.StickyLayoutHeader.isTitle > header > button.ant-btn.ant-btn-success.ant-dropdown-trigger.BTDropdown.BTButton.AutoSizing")).Click();
            d.FindElement(By.CssSelector("#rc-tabs-0-panel-1 > div > div.GridContainer-Header.StickyLayoutHeader.isTitle > header > div > div > div > ul > li:nth-child(1) > span > a")).Click();
        }

        static string AddDaysToToday(int day)
        {
            System.DateTime today = System.DateTime.Now;
            System.TimeSpan duration = new System.TimeSpan(day, 0, 0, 0);
            System.DateTime answer = today.Add(duration);
            System.Console.WriteLine(answer);
            string date = Convert.ToString(answer);
            return date;
        }

        static void InputPO(WebDriver d, IDictionary <String, String> row) {
            Thread.Sleep(7000);
            // Enter Title
            IWebElement e = d.FindElement(By.CssSelector("#title"));
            e.SendKeys(row["Title"]);
            Thread.Sleep(1000);

            //Enter Assign to
            e = d.FindElement(By.CssSelector("#performingUserId"));
            e.SendKeys(row["Assigned to"] + OpenQA.Selenium.Keys.Enter);
            /*e.SendKeys(row["Assigned to"]);
            Thread.Sleep(3000);
            e.SendKeys(OpenQA.Selenium.Keys.Enter);*/
            Thread.Sleep(3000);
            try {
                e = d.FindElement(By.XPath("//*[@id=\"ctl00_ctl00_bodyTagControl\"]/div[15]/div/div[2]/div/div[2]/div/div/div[2]/button[2] AND  ")  );
            }
            catch {
                
            }
            

            //Click the Item button
            e.SendKeys(OpenQA.Selenium.Keys.PageDown);
            e.SendKeys(OpenQA.Selenium.Keys.PageDown);
            Thread.Sleep(1000);
            //#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div.ant-modal-body > div > div.ModalContentContainer > form > main > div > div.ant-col.margin-bottom-xs.ant-col-xs-24.ant-col-sm-18 > div.ant-card.PageSection.removeBodyPadding > div > div:nth-child(5) > div.ant-card-body > div > div:nth-child(2) > form > div > div > div > div > div > div > div > div > div.ant-table-body > div > table > tbody > tr.ant-table-row.ant-table-row-level-0.actionRow.none > td.ant-table-cell.ant-table-cell-fix-left.ant-table-cell-fix-left-last.text-left > button
            //I don't know why the directory is changing, but this selector work the best (19-20-24-26-27) If start from beginning 24 works, next will be 29 (+5)
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div.ant-modal-body > div > div.ModalContentContainer > form > main > div > div.ant-col.margin-bottom-xs.ant-col-xs-24.ant-col-sm-18 > div.ant-card.PageSection.removeBodyPadding > div > div:nth-child(5) > div.ant-card-body > div > div:nth-child(2) > form > div > div > div > div > div > div > div > div > div.ant-table-body > div > table > tbody > tr.ant-table-row.ant-table-row-level-0.actionRow.none > td.ant-table-cell.ant-table-cell-fix-left.ant-table-cell-fix-left-last.text-left > button"));
            Thread.Sleep(1000);
            e.Click();
            Thread.Sleep(1000);
            
            //Send Title2
            e = d.FindElement(By.CssSelector("#purchaseOrderLineItems\\[0\\]\\.itemTitle"));
            e.SendKeys(row["Title2"]);
            Thread.Sleep(1000);
            
            //Send Unit Cost, Clear Unit Const by send 6 Backspaces
            e = d.FindElement(By.CssSelector("#purchaseOrderLineItems\\[0\\]\\.unitCost"));
            e.SendKeys(OpenQA.Selenium.Keys.Backspace); e.SendKeys(OpenQA.Selenium.Keys.Backspace); e.SendKeys(OpenQA.Selenium.Keys.Backspace); e.SendKeys(OpenQA.Selenium.Keys.Backspace); e.SendKeys(OpenQA.Selenium.Keys.Backspace); e.SendKeys(OpenQA.Selenium.Keys.Backspace);
            Thread.Sleep(1000);
            e.SendKeys(row["Unit Cost"] + OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
            
            //Send Cost Code
            e = d.FindElement(By.CssSelector("#purchaseOrderLineItems\\[0\\]\\.costCodeId"));
            e.SendKeys(row["Cost Code"] + OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

            //Click outsite item and save
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.ModalContentContainer > form > main > div > div.ant-col.ant-col-xs-24.ant-col-sm-6"));
            e.Click();
            Thread.Sleep(1000);
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div.ant-modal-body > div > div.BTModalFooter.Unstuck > button:nth-child(1)"));
            e.Click();
            Thread.Sleep(5000);

            //Grab invoice number
            e = d.FindElement(By.CssSelector("#purchaseOrderName"));
            Thread.Sleep(1000);
            string num = e.GetAttribute("value");
            Console.WriteLine("Invoive Number is:" + num);

            //Create New Payment Bill
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.ModalContentContainer > form > main > div > div.ant-col.margin-bottom-xs.ant-col-xs-24.ant-col-sm-18 > div.ant-card.PageSection.purchaseOrder-billsLienWaiversList > div.ant-card-head > div > div.ant-card-extra > a > button"));
            e.Click();
            Thread.Sleep(3000);
            
            //Click apply 100%
            //e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(29) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.ModalContentContainer > div:nth-child(2) > main > div > div.ant-card-body > div.ant-row.ant-row-bottom.BTRow-xs > div:nth-child(2) > button"));
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(24) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.ModalContentContainer > div:nth-child(2) > main > div > div.ant-card-body > div.ant-row.ant-row-bottom.BTRow-xs > div:nth-child(2) > button"));
            e.Click();
            Thread.Sleep(1000);
            
            //Click save for apply --then bump out bill window
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(24) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.BTModalFooter > button"));
            e.Click();
            Thread.Sleep(10000);
            
            //Save apply -then everyting uneditable
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(24) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.BTModalFooter.Unstuck > button:nth-child(1)"));
            e.Click();
            Thread.Sleep(10000);
            
            //Close Bill
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(24) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.BTModalHeader > button"));
            e.Click();
            Thread.Sleep(1000);
            
            //Save Purchase Order
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.BTModalFooter.Unstuck > button:nth-child(1)"));
            e.Click();
            Thread.Sleep(10000);
            
            //Close Purchase Order
            e = d.FindElement(By.CssSelector("#ctl00_ctl00_bodyTagControl > div:nth-child(19) > div > div.ant-modal-wrap.buildertrend-custom-modal.buildertrend-custom-modal-no-header > div > div.ant-modal-content > div > div > div.BTModalHeader.Unstuck > button"));
            e.Click();
            Thread.Sleep(1000);
            String projectNo = Convert.ToString(row["Project No"]) + "-" + Convert.ToString(num);
            Console.WriteLine("{0} is saved!", projectNo);
        }

        static void DeleteFromInputSheet() {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();
            ExcelRowCollection rows = worksheet.Rows;
            // Delete the 2nd row from the worksheet.
            rows.Remove(1);
            workbook.Save("Input sheet.xlsx");
        }

        static void ClearSearchBoxGoBackSummary(ChromeDriver d)
        {
            //Clear the Search Box
            Thread.Sleep(2000);
            d.FindElement(By.CssSelector("#reactJobPicker > div > div.JobPickerHeader > div.SearchContainer > span > span > button")).Click();//Clear by click x
            Thread.Sleep(5000);
            //Clear Search List
            d.FindElement(By.CssSelector("#reactJobPicker > div > div.ant-list.ant-list-split.BTListVirtual.JobList > div > div > div:nth-child(1) > div > div > li.ant-list-item.JobListItem.AllJobs > div > div"));
            Thread.Sleep(2000);
            //Go back to Summary
            var j = d.FindElement(By.CssSelector("#reactMainNavigation > div > div.MainNavDropdownsRow.darken > div > div:nth-child(2) > button"));
            j.Click();
            Thread.Sleep(1000);
            var s = d.FindElement(By.CssSelector("#reactMainNavigation > div > div.MainNavDropdownsRow.darken > div > div:nth-child(2) > div > div > div > ul > li:nth-child(1) > span > div > div > a > div"));
            s.Click();
            Thread.Sleep(2000);
        }
    }
}
