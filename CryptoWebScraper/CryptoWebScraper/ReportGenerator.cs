using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Drawing;
using System.Drawing.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CryptoWebScraper
{
    class ReportGenerator
    {
        public String Name { get; set; }
        public DateTime TimeStamp { get; set; }
        public String dateTimeStr { get; set; }
        /* Public method to generate html report from data in database */
        public void GenerateHTML()
        {
            var query = Query();
            /* Echo query result in the console and build html page for query result */
            Console.WriteLine("Getting info from db for " + Name);
            var report = new StringBuilder();
            report.AppendLine("<!DOCTYPE html ><html><head><meta charset = 'ISO-8859-1'><title>" + Name +
                              " Reports </title>");
            report.AppendLine("<style> table, th, td { border: 1px solid black; border-collapse: collapse; }");
            report.AppendLine("th, td { padding: 5px; }");
            report.AppendLine("th { text - align: left; }");
            report.AppendLine("</style></head><body><table style='width: 100 % '>");
            string[] tags = { "<tr>", "</tr>", "<td>", "</td>", "</table>", "<th>", "</th>" };
            report.AppendLine(tags[0]);
            report.AppendLine(tags[5] + "No" + tags[6]);
            report.AppendLine(tags[5] + "Name" + tags[6]);
            report.AppendLine(tags[5] + "Market Cap" + tags[6]);
            report.AppendLine(tags[5] + "Price" + tags[6]);
            report.AppendLine(tags[5] + "Volume (24h)" + tags[6]);
            report.AppendLine(tags[5] + "Circulating Supply" + tags[6]);
            report.AppendLine(tags[5] + "Change (24h)" + tags[6]);
            report.AppendLine(tags[5] + "TimeStamp" + tags[6]);
            report.AppendLine(tags[1]);
            foreach (var cc in query)
            {
                Console.WriteLine("Item:" + cc.Item + " Name:" + cc.Name + " MarketCap:" + cc.MarketCap +
                                  " Price:" +
                                  cc.Price + " Volume:" + cc.Volume + " Supply:" + cc.Supply + " Change:" + cc.Change +
                                  " TimeStamp:" +
                                  cc.TimeStamp);
                report.AppendLine(tags[0]);

                report.AppendLine(tags[2] + cc.Item + tags[3]);
                report.AppendLine(tags[2] + cc.Name + tags[3]);
                report.AppendLine(tags[2] + cc.MarketCap + tags[3]);
                report.AppendLine(tags[2] + cc.Price + tags[3]);
                report.AppendLine(tags[2] + cc.Volume + tags[3]);
                report.AppendLine(tags[2] + cc.Supply + tags[3]);
                report.AppendLine(tags[2] + cc.Change + tags[3]);
                report.AppendLine(tags[2] + cc.TimeStamp + tags[3]);

                report.AppendLine(tags[1]);
            }
            report.AppendLine(tags[4] + "</body></html>");
            //Console.WriteLine(report);
            //if (dateTimeStr == "")  Utils.getDateTimeStr(DateTime.Now);
            StreamWriter sw = new StreamWriter("C:/testdata/testhtml-" + dateTimeStr + ".html");
            sw.Write(report);
            sw.Flush();
            sw.Close();
            IWebDriver driver = new ChromeDriver();
            /* Display html page with query result */
            driver.Navigate().GoToUrl("file:///C:/testdata/testhtml-" + dateTimeStr + ".html");

        }


        public void GenerateExcel()
        {
            var query = Query();
            /* Echo query result in the console and build html page for query result */
            Console.WriteLine("Creating Excel Spreadsheet for ... " + Name);
            Excel.Application excelApp = new Excel.Application
            {
                DisplayAlerts = false
            };
#if DEBUG
            {
                //excelApp.Visible = true;
                // launch = true;

            }
#else
//excelApp.Visible = false; 
#endif
            //Get a new workbook. 
            Excel.Workbook book = (Excel.Workbook)(excelApp.Workbooks.Add(Missing.Value));
            Excel.Worksheet sheet = (Excel.Worksheet)book.ActiveSheet;
            string[] headers =
            {
                "No", "Name", "Market Cap", "Price", "Volume (24h)", "Circulating Supply", "Change (24h)", "TimeStamp"
            };
            sheet.Cells[1, 1] = headers[0];
            sheet.Cells[1, 2] = headers[1];
            sheet.Cells[1, 3] = headers[2];
            sheet.Cells[1, 4] = headers[3];
            sheet.Cells[1, 5] = headers[4];
            sheet.Cells[1, 6] = headers[5];
            sheet.Cells[1, 7] = headers[6];
            sheet.Cells[1, 8] = headers[7];
            //Format the Header row to make it Bold and blue
            sheet.get_Range("A1", "H1").Interior.Color = Color.SkyBlue;
            sheet.get_Range("A1", "H1").Font.Bold = true;
            //Set the column widthe of Column A and Column B to 20
            sheet.get_Range("A1", "H1").ColumnWidth = 20;
            int counter = 1;
            foreach (var cc in query)
            {

                sheet.Cells[counter, 1] = cc.Item;
                sheet.Cells[counter, 2] = cc.Name;
                sheet.Cells[counter, 3] = cc.MarketCap;
                sheet.Cells[counter, 4] = cc.Price;
                sheet.Cells[counter, 5] = cc.Volume;
                sheet.Cells[counter, 6] = cc.Supply;
                sheet.Cells[counter, 7] = cc.Change;
                sheet.Cells[counter, 8] = cc.TimeStamp;
                counter++;
            }
            //if (dateTimeStr == "") Utils.getDateTimeStr(DateTime.Now);
            String reportFile = "C:\\testdata\\testexcel-" + dateTimeStr + ".xlsx";


            book.SaveAs(reportFile,
                Excel.XlFileFormat.xlWorkbookDefault,
                Type.Missing,
                Type.Missing,
                false,
                false,
                Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing);



            excelApp.Quit();

            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(excelApp);

            sheet = null;
            book = null;
            excelApp = null;
            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);
            Utils.OpenExcelFile(reportFile);

        }

        private IOrderedQueryable<CryptoCurrency> Query()
        {
            /* build query for databse */
            var db = new CryptoCurrencyContext();
            var query = from crypcur in db.CryptoCurrencies
                        where crypcur.Name == Name && (crypcur.TimeStamp <= TimeStamp)
                        orderby crypcur.TimeStamp
                        select crypcur;
            /* if there is no data in the query result then print error */
            if (query.Count() < 1)
            {
                Console.WriteLine("Query Error: Name=" + Name + "TimeStamp=" + TimeStamp);
            }
            return query;
        }
    }
}
