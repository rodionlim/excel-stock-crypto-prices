using System;
using System.Threading.Tasks;
using Xunit;

namespace StockScraper.Tests
{
    public class StockScraperIntegrationTests
    {
        [Fact]
        public async Task GetHTML_ValidTicker_ReturnsNonNullData()
        {
            // arrange
            // var ticker = "ES3.SI";
            var ticker = "O39.SI";

            // act
            var result = await Scraper.GetHTML(ticker);

            // assert
            Assert.NotNull(result);
        }
    }
}
