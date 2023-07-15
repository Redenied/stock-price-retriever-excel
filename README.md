# stock-price-retriever-excel
Automated stock price retriever for Excel file in VBA using Selenium

This VBA script will help any stock investor automatically update their choice stock prices with the latest price according to yahoo finance.

### Requirements:
- Microsoft Excel
- Selenium WebDriver
- Chrome web browser

### Installation:
- Clone this repository
- Install Selenium WebDriver. I downloaded it from https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0 (copyrights florentbr)
- Download the ChromeDriver version that is compatible with your Chrome browser https://chromedriver.chromium.org/downloads
- To find out which Chrome version you own: Go to Chrome settings > Help > About Google Chrome
- Copy the ChromeDriver executable 'chromedriver.exe' you downloaded and paste it to your local directory \AppData\Local\SeleniumBasic and replace the existing executable

### Usage:
1. Open the Excel worksheet
2. Place in column 'C' the symbols of the stocks you want to get the share prices. Be careful to write the exact name that you can find on Yahoo Finance.
3. Fill columns 'D' and 'E' with your average prices in order to calculate the profit.
4. Fill cell 'D32' with the cash value of your portfolio, to calculate the weight of each portfolio element.

Be careful to not modify column 'K', you can hide it if you want. That's where the macro gets the stock link from.
Right now the code works with 28 stocks, if you want to add more you need to modify the VBA script where the For loop is and modify the range values.

Thank you, any submissions are welcome!
Redenied
