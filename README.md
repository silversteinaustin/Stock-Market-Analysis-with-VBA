# VBA Challenge - Stock Market Analysis

## Overview

In the quest to become a proficient programmer and Excel expert, this project utilizes VBA scripting to analyze stock market data across multiple years. The goal is to automate the analysis process, making it efficient to extract key metrics about stock performance directly within Excel. This analysis provides insights into yearly stock performance, including changes, percentage changes, and total stock volumes.

### Objective

The primary objective of this VBA Challenge is to create a robust VBA script capable of looping through each year of stock data to output critical stock performance indicators. Additionally, the script identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume, facilitating strategic investment decisions.

### Methodology

#### Data Retrieval

- **Ticker Symbol:** Extracted and stored from each row.
- **Volume of Stock:** Calculated the total stock volume for each ticker.
- **Open and Close Prices:** Identified to calculate yearly changes and percentage changes.

#### Analysis Outputs

- **Yearly Change:** Calculated from the opening price at the beginning of the year to the closing price at the end of that year.
- **Percentage Change:** Derived from the yearly change in relation to the opening price.
- **Total Stock Volume:** Aggregated for each stock to assess trading activity.

#### Additional Functionality

- **Greatest % Increase, Decrease, and Total Volume:** Script identifies and outputs these key metrics across all stocks analyzed.
- **Multi-Year Compatibility:** Adjusted the VBA script to operate across every worksheet, allowing for a comprehensive analysis with a single script execution.

### Key Findings

- The analysis successfully extracted and calculated the required metrics for stock performance evaluation.
- Conditional formatting was applied to visually distinguish between positive and negative yearly changes, enhancing the readability of results.


### Conclusion

This project demonstrates the power of VBA scripting in automating data analysis within Excel, significantly reducing manual effort and potential for error. While there were challenges in ensuring the accuracy of certain metrics, the overall project was a valuable exercise in applying programming skills to practical financial data analysis.

### Acknowledgments

I appreciate the opportunity to tackle this VBA Challenge and thank the reviewers for their constructive feedback, which was instrumental in refining the project's outcomes.
