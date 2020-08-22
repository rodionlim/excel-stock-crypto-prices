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
Depending on whether Excel is 32 or 64bit, choose root dir's StockScraper-AddIn-packed*.xll file
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

To get the latest closing stock price:
```
=bdp("ES3.SI")
```
