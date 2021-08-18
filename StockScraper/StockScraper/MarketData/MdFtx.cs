using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;

using Newtonsoft.Json;

namespace StockScraper
{
    class MdFtx
    {
        private static readonly HttpClient _client = new HttpClient();
        public static async Task<Dictionary<string, string>> GetQuoteSummary(string ticker)
        {
            // ftx ticker: btc-0924, btc/usd
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpResponseMessage response = await _client.GetAsync($"https://ftx.com/api/markets/{ticker}");

            // Read the content.
            string responseFromServer = await response.Content.ReadAsStringAsync();
    
            // Augment the content.
            dynamic obj = JsonConvert.DeserializeObject(responseFromServer);
            Dictionary<string, string> quoteSummary;
            if (obj.success == true)
            {
                var result = obj.result;
                quoteSummary = result.ToObject<Dictionary<string, string>>();
                quoteSummary.Add("CURRENT PRICE", quoteSummary["price"]);
            }
            else
                throw new ArgumentException(obj["error"]);
            return quoteSummary;
        }
    }
}
