using AngleSharp;
using AngleSharp.Io;
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
			"52-WEEK RANGE",
			"VOLUME",
			"AVG. VOLUME",
			"NET ASSETS",
			"NAV",
			"PE RATIO (TTM)",
			"YIELD",
			"YTD DAILY TOTAL RETURN",
			"BETA (5Y MONTHLY)",
			"EXPENSE RATIO (NET)",
			"INCEPTION DATE",
			"MARKET CAP",
			"EPS (TTM)",
			"EARNINGS DATE",
			"FORWARD DIVIDEND & YIELD",
			"EX-DIVIDEND DATE",
			"1Y TARGET EST"
		};

		[ExcelAsyncFunction(Description = "Grab latest Yahoo/Google finance market data of a single exchange traded ticker, such as the closing price, dividend yield etc., defaults to Yahoo finance.")]
		public static async Task<object> bdp(
			[ExcelArgument(Name = "ticker", Description = "Yahoo/Google finance ticker. Example: \"es3.si\", \"goog\"")] string ticker,
			[ExcelArgument(AllowReference = true, Name = "[field]", Description = "[Optional] Yahoo finance or Google fields. Example: \"previous close\", \"market cap\", \"yield\"")] string field,
			[ExcelArgument(AllowReference = true, Name = "[source]", Description = "[Optional] Specify the Data Source, defaults to yahoo. Example: \"yahoo\", \"google\"")] string source

		)
		{
            var defaultField = "CURRENT PRICE";

			field = field.ToUpper();
			if (field.Length == 0) field = defaultField;
			if (!AvailableFields.Contains(field)) return "#Invalid field"; // Validate user input

			if (source.Length==0) source = "yahoo";
			if (!(source == "yahoo" | source == "google")) return "#Invalid data source"; // Validate user input

			if (source == "google" & field != defaultField) return "#Google finance currently only supports current price"; // Validate user input

			Dictionary<string, string> quoteSummary;
			if (source == "yahoo")
				quoteSummary = await GetHTML(ticker);
			else
				quoteSummary = await GetHTMLGoogle(ticker);

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

		private static async Task<Dictionary<string, string>> GetHTMLGoogle(string ticker)
		{
			var quoteSummary = new Dictionary<string, string>();
			var requester = new DefaultHttpRequester();
			requester.Headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36";
			var config = Configuration.Default.With(requester).WithDefaultLoader();
			var context = BrowsingContext.New(config);
			var document = await context.OpenAsync($"https://www.google.com/search?q={ticker}");

			var quoteSummaryListItems = document.QuerySelector("span[jsname='vWLAgc']");

			var tmp = quoteSummaryListItems.TextContent.Replace(",", "");
			quoteSummary.Add("CURRENT PRICE", tmp);
			return quoteSummary;
		}
	}
}
