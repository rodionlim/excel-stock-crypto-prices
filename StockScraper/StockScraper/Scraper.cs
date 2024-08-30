using AngleSharp;
using AngleSharp.Io;
using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Net.Http;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace StockScraper
{
	public class Scraper
	{
		private static readonly List<string> AVAILABLE_FIELDS = new List<string>()
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

		private static readonly string USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36";

		[ExcelAsyncFunction(Description = "Grab latest Yahoo/Google finance or FTX market data of a single exchange traded ticker, such as the closing price, dividend yield etc., defaults to Yahoo finance and Current Price.")]
		public static async Task<object> bdp(
			[ExcelArgument(Name = "ticker", Description = "Yahoo/Google finance or FTX ticker. Example: \"es3.si\", \"goog\", \"btc/usd\", \"btc-0924\"")] string ticker,
			[ExcelArgument(AllowReference = true, Name = "[field]", Description = "[Optional] Yahoo finance or Google fields. Example: \"previous close\", \"market cap\", \"yield\"")] string field,
			[ExcelArgument(AllowReference = true, Name = "[source]", Description = "[Optional] Specify the Data Source, defaults to yahoo. Example: \"yahoo\", \"google\", \"ftx\"")] string source

		)
		{
            var defaultField = "CURRENT PRICE";

			field = field.ToUpper();
			if (field.Length == 0) field = defaultField;
			if (!AVAILABLE_FIELDS.Contains(field)) return "#Invalid field"; // Validate user input

			if (source.Length==0) source = "yahoo";
			if (!(source == "yahoo" | source == "google" | source =="ftx")) return "#Invalid data source"; // Validate user input

			if ((source == "google" | source == "ftx") & field != defaultField) return "#Google finance and FTX data sources currently only supports current price"; // Validate user input

			Dictionary<string, string> quoteSummary;
			switch (source)
			{
				case "yahoo":
					quoteSummary = await GetHTML(ticker);
					break;
				case "google":
					quoteSummary = await GetHTMLGoogle(ticker);
					break;
				case "ftx":
					try	{ quoteSummary = await MdFtx.GetQuoteSummary(ticker);}
					catch (ArgumentException e) { return e.Message; }
					break;
				default:
					return "#Invalid data source";
			}

			if (double.TryParse(quoteSummary[field], out double res))
			{
				return res;
			}
			return quoteSummary[field];
		}

		internal static async Task<Dictionary<string, string>> GetHTML(string ticker)
		{
			var quoteSummary = new Dictionary<string, string>();
			var htmlRequestor = new DefaultHttpRequester();
			htmlRequestor.Headers["User-Agent"] = USER_AGENT;
			var config = Configuration.Default.With(htmlRequestor).WithDefaultLoader();
			var context = BrowsingContext.New(config);
			var tickerUC = ticker.ToUpper();
			var document = await context.OpenAsync($"https://sg.finance.yahoo.com/quote/{tickerUC}/");

			var quoteSummaryListItems = document.QuerySelectorAll("div[data-testid='quote-statistics'] li");
			var quoteHeaderInfo = document.QuerySelectorAll($"fin-streamer[data-field='regularMarketPrice'][data-symbol='{tickerUC}']");
			if (quoteHeaderInfo.Length >= 1)
			{
				var tmp = quoteHeaderInfo[0].TextContent;
                if (float.TryParse(tmp, out float currentPx))
                    quoteSummary.Add("CURRENT PRICE", tmp);
            }
			else
				quoteSummary.Add("CURRENT PRICE", "#NA");

			foreach (var item in quoteSummaryListItems)
			{
				quoteSummary.Add(item.Children[0].TextContent.ToUpper(), item.Children[1].TextContent);
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
