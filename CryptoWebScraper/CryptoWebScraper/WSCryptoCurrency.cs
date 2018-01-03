using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CryptoWebScraper
{
    class WSCryptoCurrency
    {
        public void Execute()
        {
            /* Set-up for using Chrome for to browse wesite */
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://coinmarketcap.com/");

            /* Get reference to Currency table with unique identifier */
            IWebElement table = driver.FindElement(By.Id("currencies"));
            /* Get a collection tr tags in the table */
            List<IWebElement> rows = table.FindElements(By.TagName("tr")).ToList();
            //Console.WriteLine("rows: "  + rows.Count);

            /* Collection of CryptoCurrency  */
            var cryptocurrencies = new List<CryptoCurrency>();
            CryptoCurrency cryptocurrency;
            /* To capture and store date and time when scraping was done */
            DateTime timestamp = DateTime.Now;
            /* To setup delimiter to be used for splitting text into words */
            char[] delimiterChars = { ' ' };
            /* Go thru each table row(tr's)  and then process the collection of td's  for that row */
            foreach (var row in rows)
            {

                /* Get collection of td's for each  tr */
                List<IWebElement> tds = row.FindElements(By.TagName("td")).ToList();
                //Console.WriteLine("tds: " + rows.Count);
                /* Setup counter to track data in a row */
                int tdCounter = 1;
                cryptocurrency = null;
                bool finish = false;
                /* Go thru each collection of td's */
                foreach (var td in tds)
                {  /* if first data element in a row create object for CryptoCurrency */
                    if (tdCounter == 1) cryptocurrency = new CryptoCurrency();
                    //Console.WriteLine(td.Text);
                    /* Populate CryptoCurrency object with data from a row */
                    switch (tdCounter)
                    {
                        case 1:
                            {

                                int itemNo;
                                /* Convert String data to int without triggering an exception */
                                if (Int32.TryParse(td.Text, out itemNo))
                                {
                                    cryptocurrency.Item = itemNo;

                                }
                                else
                                {
                                    Console.WriteLine("ItemNo Conversion Failed");
                                }
                                break;
                            }
                        case 2:
                            {
                                cryptocurrency.Name = td.Text;
                                break;
                            }
                        case 3:
                            {
                                cryptocurrency.MarketCap = td.Text;
                                break;
                            }
                        case 4:
                            {
                                cryptocurrency.Price = td.Text;
                                break;
                            }
                        case 5:
                            {
                                cryptocurrency.Volume = td.Text;
                                break;
                            }
                        case 6:
                            {
                                /* Splitting data into separate words and using only word required */
                                string[] words = td.Text.Split(delimiterChars);
                                cryptocurrency.Supply = words[0];
                                break;
                            }
                        case 7:
                            {
                                cryptocurrency.Change = td.Text; ;
                                break;
                            }
                        case 8:
                            {
                                /* Use  common datetime for all Currenies during scraping session */
                                cryptocurrency.TimeStamp = timestamp;
                                /* Last item completed; mark object ready for collection */
                                finish = true;
                                break;
                            }
                        default:
                            {
                                /* catch all for data you are not interested in */
                                //System.Console.WriteLine("Other number");
                                break;
                            }
                    }

                    tdCounter++;
                }
                /* Move Currency object to collection */
                if (finish)
                    cryptocurrencies.Add(cryptocurrency);
            }
            var db = new CryptoCurrencyContext();
            /* Go thru Curreny Collection and echo to console and populate database */
            foreach (var cc in cryptocurrencies)
            {
                //Console.WriteLine("Item:"+ cc.Item + " Name:" + cc.Name + " MarketCap:" + cc.MarketCap + " Price:" + cc.Price );
                Console.WriteLine("Item:" + cc.Item + " Name:" + cc.Name + " MarketCap:" + cc.MarketCap + " Price:" + cc.Price + " Volume:" + cc.Volume + " Supply:" + cc.Supply + " Change:" + cc.Change + " TimeStamp:" + cc.TimeStamp);
                /* Populate database */
                AddCryptoCurrency(db, cc);

            }
            /* Generate a report from the database base on the name of the currency and a datetime value */
            GetCryptoCurrency("Bitcoin", DateTime.Now);
        }
        /* Database writing of scraped data */
        private static void AddCryptoCurrency(CryptoCurrencyContext db, CryptoCurrency cryptoCurrency)
        {
            db.CryptoCurrencies.Add(cryptoCurrency);
            db.SaveChanges();
        }
        /* Generating an html report from currency data in the database */
        private static void GetCryptoCurrency(String name, DateTime datetime)
        {
            var reportgenerator = new ReportGenerator
            {
                Name = name,
                TimeStamp = datetime,
                dateTimeStr = Utils.getDateTimeStr(DateTime.Now)
            };
            reportgenerator.GenerateHTML();
            reportgenerator.GenerateExcel();

        }
    }
}
