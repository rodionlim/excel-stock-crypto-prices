# Excel Stock Price Extraction
This repository provides an excel addin with user defined functions to scrape stock market data from Yahoo Finance

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

To get the latest market data for a exchange traded ticker:
```
Syntax: =bdp( {ticker}, {yahoo field} )

=bdp("ES3.SI") - STI ETF's previous closing prrice
=bdp("AAPL", "yield") - Apple's dividend yield
=bdp("MSFT", "volume") - Microsoft's volume traded
```

List of available fields
```
PREVIOUS CLOSE, 
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
