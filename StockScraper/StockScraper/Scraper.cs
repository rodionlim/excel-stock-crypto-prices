using AngleSharp;
using AngleSharp.Text;
using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace StockScraper
{
    public static class Scraper
	{
		[ExcelAsyncFunction(Description = "Grab latest closing price of a single ticker")]
		public static async Task<double> bdp(
			[ExcelArgument(Name = "ticker", Description = "Yahoo finance ticker. Example: \"es3.si\"")] string ticker
		)
		{
			var quoteSummary = await GetHTML(ticker);
			return quoteSummary["Previous close"].ToDouble();
		}

		private static async Task<Dictionary<string, string>> GetHTML(string ticker)
		{
			var quoteSummary = new Dictionary<string, string>();
			string keyCache = "";
			var config = Configuration.Default.WithDefaultLoader();
			var context = BrowsingContext.New(config);
			var document = await context.OpenAsync($"https://sg.finance.yahoo.com/quote/{ticker}/");

			var quoteSummaryListItems = document.QuerySelectorAll("div[id='quote-summary'] td");

			foreach (var item in quoteSummaryListItems.Select((sli, i) => new { Value = sli, Index = i }))
			{
				if (item.Index % 2 == 0)
					keyCache = item.Value.TextContent;
				else
					quoteSummary.Add(keyCache, item.Value.TextContent);
			}

			return quoteSummary;
		}
	}
}
