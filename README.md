# VBA-challenge

The purpose of this project is to automate the analysis of large datasets of quarterly data for thousands of stock ticker symbols, easily saving hours upon hours of repetitive manual work.  

The code created for this project provides:

- A summary of all unique ticker symbols provided along with
    - The current quarter's net change (calculated by using the current quarter's closing price net of the current quarter's opening price),
        along with an application of conditional formatting to highlight gains/losses for the quarter
    - The current quarter's percentage change
    - A summation of the current quarter's volume of total shares traded
    
- A quarterly highlights chart which provides the stock tickers that have had the
    1) greatest percentage increase for the period,
    2) greatest percentage decrease for the period, and
    3) greatest volume of shares traded. 

This code maintains consistent and dependable outputs, regardless of the number of lines of data for a given quarter, and is able to run across multiple worksheets within a given workbook.

Limitations: The code was written with the assumption that the data provided will always be provided sorted by ticker symbol, then by trade date. Also, the data provided will consist of a full quarter's data per worksheet.

Any questions? 

Feel free to send a message to acdlc4@gmail.com with any questions / comments / concerns.

Inspiration and credit for any code used was obtained from work done during my attendance in 2024 Northwestern University Data Analysis Bootcamp class sessions.
