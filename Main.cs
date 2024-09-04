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
using GemBox.Spreadsheet.Tables;
using System.Xml.Linq;
using System.Linq.Expressions;
using SkiaSharp;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.AxHost;
using System.Text.RegularExpressions;

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
                SearchNewPO(Login, r);
                Console.WriteLine("**********************************");
                foreach (KeyValuePair<string, string> ele in r)
                    Console.WriteLine("{0}: {1}", ele.Key, ele.Value);
                if (!CheckDuplicatePO(Login, r))
                {
                    StartNewPO(Login);
                    InputPO(Login, r);
                }
                else { Console.WriteLine("--Exsixting item--"); }

                DeleteFromInputSheet();
                ClearSearchBox(Login);
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
            try
            {
                var b = d.FindElement(By.XPath("//html/body/div[2]/div/div/div[3]/form/div[3]/div[4]/div/div/div[1]/div/div[1]/div/div[6]/button"));
                b.Click();
                Thread.Sleep(2000);
                //Find Purchase Order
                var BP = d.FindElement(By.XPath("//*[@id=\"reactMainNavigation\"]/div/div[1]/div/div[6]/div/div/div/ul/li[3]/span/div/div/a/div/div/div[2]/div"));
                BP.Click();
            }
            catch
            {
                Console.WriteLine("Failed to goto Purchase Order Page. Please Manually goto Purchase Order Page and leave the Chrome visible.");
            }
            Thread.Sleep(3000);

            //Clear Button
            try {
                var b = d.FindElement(By.XPath("//*[@data-testid='clear-search' and @type='button']"));
                b.Click();
                Thread.Sleep(10000);
            }
            catch {
            }

            // Close the chatbox if possible
            try
            {
                //IFrame - Close ChatBox
                //Switch to the frame
                d.SwitchTo().Frame("intercom-launcher-frame");
                Thread.Sleep(3000);
                //Now click the button
                var e = d.FindElement(By.XPath("//*[@id='intercom-container']/div/div/div[2]"));
                e.Click();
                Thread.Sleep(1000);
                // Return to the top level
                d.SwitchTo().DefaultContent();
                Thread.Sleep(3000);
                //Click Close Button
                e = d.FindElement(By.CssSelector("#btnCloseIntercom"));
                e.Click();
                Thread.Sleep(1000);
            }
            catch
            {
                //Console.WriteLine("No Chat Box");
            }
            finally
            {
                d.SwitchTo().DefaultContent();
            }

            // Close Updates From BuilderTrends
            try
            {
                //IFrame - Close ChatBox
                //Switch to the frame
                d.SwitchTo().Frame("intercom-notifications-frame");
                Thread.Sleep(3000);
                //Now click the button
                var e = d.FindElement(By.XPath("//*[@id=\"intercom-container\"]/div/div/div/div/div/div[1]/button"));
                e.Click();
                Thread.Sleep(1000);
                // Return to the top level
                d.SwitchTo().DefaultContent();
                Thread.Sleep(3000);
                }
            catch
            {
                //Console.WriteLine("No Updates From BuilderTrends");
            }
            finally
            {
                d.SwitchTo().DefaultContent();
            }

        }

        static int CheckProjectsNum()
        {
            ExcelFile workbook = ExcelFile.Load("Input sheet.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets.First();
            int rows = worksheet.Rows.Count();
            return rows;
        }

        static void SearchNewPO(ChromeDriver d, IDictionary<String, String> row)
        {
            try
            {
                Thread.Sleep(2000);
                IWebElement e = d.FindElement(By.Id("JobSearch"));
                e.SendKeys(row["Project No"]);
                Thread.Sleep(2000);
                e = d.FindElement(By.ClassName("ItemRowJobName"));
                e.Click();// Click to Job Order
                Thread.Sleep(5000);
            }
            catch {
                Console.WriteLine("Failed to Search PO number.");
            }
        }
        

        static bool CheckDuplicatePO(WebDriver d, IDictionary<String, String> row)
        {
            // Clean the project invioce number into digits
            string s = row["Title"];
            string symbleNumber = Regex.Replace(s, @"([a-zA-Z]+)", "");
            string spaceNumber = Regex.Replace(symbleNumber, @"('+)", "");
            string pureNumber = Regex.Replace(spaceNumber, @"(#+)", "");
            string invoiceDigit = pureNumber.Trim();

            try
            {
                IWebElement temp = d.FindElement(By.XPath(string.Concat("//span[contains(.,", invoiceDigit, ")]")));
                Console.WriteLine(s+ " is found duplicate");
                Thread.Sleep(5000);
                return true; 
            }
            catch {
                Thread.Sleep(5000);
                return false; 
            }
        }
        static void StartNewPO(ChromeDriver d)
        {
            //try
            //{
                //Find and click New -> PO
                IWebElement e = d.FindElement(By.XPath("//*[@data-testid='newPurchaseOrderGroup' and @type='button']"));
                e.Click();
                e = d.FindElement(By.XPath("//*[@data-testid='newPurchaseOrder' and @type='button']"));
                e.Click();
                Thread.Sleep(5000);
            //}
            //catch
            //{
                //Console.WriteLine("Failed to Search PO number.");
            //}
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

        static void InputPO(WebDriver d, IDictionary<String, String> row)
        {
                Thread.Sleep(7000);
                // Enter Title
                IWebElement e = d.FindElement(By.CssSelector("#title"));
                e.SendKeys(row["Title"]);
                Thread.Sleep(1000);

                //Enter Assign to
                e = d.FindElement(By.CssSelector("#performingUserId"));
                e.SendKeys(row["Assigned to"] + OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                // If new assignee, then Add to Job
                try
                {
                    e = d.FindElement(By.XPath("//*[@data-testid='confirmPrompt' and @type='button']"));
                    e.Click();
                    Thread.Sleep(3000);
                    Console.WriteLine("Add New Assignee to Job");
                    //Permission Wizard Update
                    try
                    {
                        e = d.FindElement(By.XPath("//*[@data-testid='permisisonWizardUpdate' and @type='button']"));
                        e.Click();
                        Thread.Sleep(1000);
                        Console.WriteLine("Updated Permission");
                    }
                    catch {
                    }
                }
                catch
                {
                }
                finally
                {
                    e = d.FindElement(By.CssSelector("#title"));
                }
           
                //Click the Item button
                e.SendKeys(OpenQA.Selenium.Keys.PageDown);
                e.SendKeys(OpenQA.Selenium.Keys.PageDown);
                Thread.Sleep(1000);
                e = d.FindElement(By.XPath("//*[text()='Item']"));
                e.Click();
                Thread.Sleep(2000);

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
                e = d.FindElement(By.XPath("//*[text()='Please Save the Purchase Order before viewing Bills.']"));
                e.Click();
                Thread.Sleep(1000);
                e = d.FindElement(By.XPath("//*[text()='Save']"));
                e.Click();
                Thread.Sleep(5000);

                //Grab invoice number
                e = d.FindElement(By.CssSelector("#purchaseOrderName"));
                Thread.Sleep(1000);
                string num = e.GetAttribute("value");
                Console.WriteLine("Invoice Number is:" + num);

                //Create New Payment Bill
                e = d.FindElement(By.XPath("//*[text()='New Bill']"));
                e.Click();
                Thread.Sleep(3000);

                //Click apply 100%
                e = d.FindElement(By.XPath("//*[text()='Apply']"));
                e.Click();
                Thread.Sleep(1000);

                //Click save for apply -- then bump out bill window
                e = d.FindElement(By.XPath("//*[@type='submit'][@data-testid='save']"));
                e.Click();
                Thread.Sleep(5000);

                //Find invoice date
                e = d.FindElement(By.XPath("//*[@id=\"invoiceDate\"]"));
                e.SendKeys(OpenQA.Selenium.Keys.Control + 'a');
                e.SendKeys(OpenQA.Selenium.Keys.Delete);
                string invoiceDate = row["Invoice Date"];
                invoiceDate = invoiceDate.Trim();
                int foundS = invoiceDate.IndexOf(" ");
                invoiceDate = invoiceDate.Remove(foundS + 1);
                e.SendKeys(invoiceDate);
                Thread.Sleep(1000);
                e.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                //Save Bill
                e = d.FindElement(By.XPath("//*[@data-testid='obpMarkReadyForPayment']/preceding-sibling::button[@data-testid='save']"));
                e.Click();
                Thread.Sleep(10000);

                //Close Bill
                e = d.FindElement(By.XPath("//*[@class='BTModalHeader']/child::button[@data-testid='close']"));
                
                e.Click();
                Thread.Sleep(5000);

                //Save Purchase Order
                e = d.FindElement(By.XPath("//*[@data-testid='saveAndNew'][@type='button']/preceding-sibling::button[@data-testid='save']"));
                e.Click();
                Thread.Sleep(10000);

                //Close Purchase Order
                e = d.FindElement(By.XPath("//*[@data-testid='close'][@type='button']"));
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

        static void ClearSearchBox(ChromeDriver d)
        {
            //Clear the Search Box
            Thread.Sleep(2000);
            d.FindElement(By.CssSelector("#reactJobPicker > div > div.JobPickerHeader > div.SearchContainer > span > span > button")).Click();//Clear by click x
            Thread.Sleep(5000);
            //Clear Search List
            d.FindElement(By.CssSelector("#reactJobPicker > div > div.ant-list.ant-list-split.BTListVirtual.JobList > div > div > div:nth-child(1) > div > div > li.ant-list-item.JobListItem.AllJobs > div > div"));
            Thread.Sleep(8000);
            //Go back to Summary
            d.Navigate().GoToUrl("https://buildertrend.net/summaryGrid.aspx");
            Thread.Sleep(2000);
        }
    }
}
