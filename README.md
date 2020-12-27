# Excel Stock Price Extraction
This repository provides an excel addin with user defined functions to scrape stock market data from Yahoo and Google Finance.  

We only support current price for Google Finance data source while the full list of fields supported for Yahoo Finance can be found below.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

What things you need to build the software (Not required if taking the dll files directly) and how to install them

- Visual Studio 2019


### Installing

A step by step series of examples that tell you how to get a development env running

```
Open Excel > File > Options > Add-ins > Manage Excel Add-ins > Browse > 
Depending on whether Excel is 32 or 64bit, choose StockScraper-AddIn*-packed.xll file
Pre-built binaries can be found in StockScraper/StockScraper/bin/Debug/ for users to use directly
```

To build the project:
```
Build the project with the .sln file
```

Link the generated xll file
```
Follow the first step
```

### Usage

Fields in brackets are optional and does not need to be specified.
To get the latest market data for an exchange traded ticker:
```
Syntax: =bdp( {ticker}, [{yahoo field}], [{data source}] )

Yahoo Finance Data Source
=bdp("ES3.SI") - STI ETF's current price
=bdp("AAPL", "yield") - Apple's dividend yield
=bdp("AAPL", "current price") - Apple's current price
=bdp("MSFT", "volume") - Microsoft's volume traded

Google Finance Data Source
=bdp("goog",,"google") - Google's current price 
=bdp("aapl",,"google") - Apple's current price
```

List of available fields
```
PREVIOUS CLOSE,
CURRENT PRICE,
OPEN,
BID,
ASK,
DAY'S RANGE,
52-WEEK-RANGE,
VOLUME,
AVG. VOLUME,
NET ASSETS,
NAV,
PE RATIO (TTM),
YIELD,
YTD DAILY TOTAL RETURN,
BETA (5Y MONTHLY),
EXPENSE RATIO (NET),
INCEPTION DATE
```
