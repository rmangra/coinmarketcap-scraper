using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CryptoWebScraper
{
    class Utils
    {
        private static Excel.Workbooks books;
        private static Excel.Workbook sheet;
        private static Excel.Application excelApp;
        public static String getDateTimeStr(DateTime datetime)
        {
            char[] delimiterChars = { ':' };
            string[] words = datetime.ToString("s").Split(delimiterChars);
            string result = String.Join("_", words);
            string[] wordmore = result.Split(delimiterChars);
            string resultmore = String.Join("_", wordmore);

            Console.WriteLine("DT:  " + resultmore);
            return resultmore;
        }

        public static void OpenExcelFile(String filename)
        {

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                books = excelApp.Workbooks;
                sheet = books.Open(filename); ;

            }
            catch (Exception e)
            {
                Console.WriteLine("Errors: " + e);
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                // Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(sheet);

                //close and release
                books.Close();
                Marshal.ReleaseComObject(books);

                //quit and release
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            finally
            {

            }
        }
    }
}
