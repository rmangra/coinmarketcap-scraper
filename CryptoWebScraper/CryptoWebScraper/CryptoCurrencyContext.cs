using System.Data.Entity;

namespace CryptoWebScraper
{
    public class CryptoCurrencyContext : DbContext
    {
        public DbSet<CryptoCurrency> CryptoCurrencies { get; set; }
    }
}
