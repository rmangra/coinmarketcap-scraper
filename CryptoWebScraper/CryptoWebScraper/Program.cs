using System;

namespace CryptoWebScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            new WSCryptoCurrency().Execute(); /* Pass control to Class to do scraping */
        }
    }
}
