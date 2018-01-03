using System;
using System.ComponentModel.DataAnnotations;

namespace CryptoWebScraper
{
    public class CryptoCurrency
    {
        [Key]
        public int Code { get; set; }
        public int Item { get; set; }
        public String Name { get; set; }
        public String MarketCap { get; set; }
        public String Price { get; set; }
        public String Volume { get; set; }
        public String Supply { get; set; }
        public String Change { get; set; }
        public DateTime TimeStamp { get; set; }
    }
}
