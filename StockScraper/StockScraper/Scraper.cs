using AngleSharp;
using ExcelDna.Integration;
using ExcelDna.Registration;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace StockScraper
{
    public class Scraper
	{
		private static readonly List<string> AvailableFields = new List<string>() 
		{
			"PREVIOUS CLOSE",
			"CURRENT PRICE",
			"OPEN",
			"BID",
			"ASK",
			"DAY'S RANGE",
			"52-WEEK-RANGE",
			"VOLUME",
			"AVG. VOLUME",
			"NET ASSETS",
			"NAV",
			"PE RATIO (TTM)",
			"YIELD",
			"YTD DAILY TOTAL RETURN",
			"BETA (5Y MONTHLY)",
			"EXPENSE RATIO (NET)",
			"INCEPTION DATE"
		};

		[ExcelAsyncFunction(Description = "Grab latest market data of a single exchange traded ticker, such as the closing price, dividend yield etc.")]
		public static async Task<object> bdp(
			[ExcelArgument(Name = "ticker", Description = "Yahoo finance ticker. Example: \"es3.si\"")] string ticker,
			[ExcelArgument(AllowReference = true, Name = "[field]", Description = "[Optional] Yahoo finance field. Example: \"previous close\", \"yield\"")] string field
		)
		{
			field = field.ToUpper();
			if (field == "") field = "PREVIOUS CLOSE";
			if (!AvailableFields.Contains(field)) return "#Invalid field"; // Validate user input

			var quoteSummary = await GetHTML(ticker);

			if (double.TryParse(quoteSummary[field], out double res))
			{
				return res;
			}
			return quoteSummary[field];
		}

		private static async Task<Dictionary<string, string>> GetHTML(string ticker)
		{
			var quoteSummary = new Dictionary<string, string>();
			string keyCache = "";
			var config = Configuration.Default.WithDefaultLoader();
			var context = BrowsingContext.New(config);
			var document = await context.OpenAsync($"https://sg.finance.yahoo.com/quote/{ticker}/");
			float currentPx;

			var quoteSummaryListItems = document.QuerySelectorAll("div[id='quote-summary'] td");
			var quoteHeaderInfo = document.QuerySelectorAll("div[id='quote-header-info'] div span");

			foreach (var item in quoteSummaryListItems.Select((sli, i) => new { Value = sli, Index = i }))
			{
				if (item.Index % 2 == 0)
					keyCache = item.Value.TextContent.ToUpper();
				else
					quoteSummary.Add(keyCache, item.Value.TextContent);
			}

			foreach (var item in quoteHeaderInfo.Select((sli) => new { Value = sli }))
			{
				// TODO: We should find a better identifier for current price instead if possible
				var tmp = item.Value.InnerHtml;
				if (float.TryParse(tmp, out currentPx))
				{
					quoteSummary.Add("CURRENT PRICE", tmp);
					break;
				}
			}

			return quoteSummary;
		}
	}
}
