# VBA-challenge
Note: This project was developed as part of the University of Toronto's Data Analytics Bootcamp program.

The Excel file containing the multi-year stock data can be found in the 'master' branch of this repository. The file is named 'Multiple_year_stock_data.xlsm'.


This VBA script is designed to efficiently analyze and provide valuable insights into multi-year stock data stored in Excel worksheets. It processes data from each worksheet and generates the following key information for each stock:

Ticker Symbol: It identifies the unique stock ticker for each entry.
Yearly Change: This represents the difference between the closing price at the end of the year and the opening price at the beginning of that year.
Percent Change: It calculates the percentage change from the opening price to the closing price, providing insights into the stock's performance.
Total Stock Volume: This measures the total volume of the stock traded throughout the year, indicating its market activity.
The VBA script employs conditional formatting to highlight positive changes in green and negative changes in red, making it easy to visualize the performance of each stock.

Moreover, the script goes beyond individual worksheets to handle multiple years of stock data. It dynamically processes and analyzes data from any worksheets with the names "2018," "2019," or "2020," ensuring that you can analyze an extensive dataset efficiently.

In addition, the script identifies and returns the stocks with the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" in the provided data. These insights can be invaluable for making investment decisions based on historical stock performance.
