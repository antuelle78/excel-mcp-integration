# Excel MCP Integration Test Prompts

## Basic Excel File Creation Tests

### Test 1: Simple Data Creation
```
Create an Excel file named "employees.xlsx" with headers ["Name", "Department", "Salary"] and data:
- John Smith, Engineering, 75000
- Jane Doe, Marketing, 65000
- Bob Johnson, Sales, 55000
```

### Test 2: Headers Only (Edge Case)
```
Create an Excel file named "structure.xlsx" with headers ["Product", "Category", "Price"] but no data rows yet.
```

### Test 3: Large Dataset
```
Create an Excel file named "sales_data.xlsx" with headers ["Date", "Product", "Quantity", "Revenue"] and include sample data for 10 sales transactions.
```

## Formatting and Styling Tests

### Test 4: Basic Formatting
```
Create an Excel file named "styled_report.xlsx" with headers ["Month", "Sales", "Expenses"] and data for 6 months. Apply bold headers with blue background.
```

### Test 5: Advanced Formatting
```
Create an Excel file named "formatted_data.xlsx" with headers ["Name", "Score", "Status"] and data. Apply:
- Bold headers with light blue background
- Alternate row colors
- Center alignment for all cells
```

## Chart Creation Tests

### Test 6: Bar Chart
```
Create an Excel file named "quarterly_sales.xlsx" with sales data for Q1-Q4, then add a bar chart showing the sales trends.
```

### Test 7: Line Chart
```
Create an Excel file named "stock_prices.xlsx" with daily stock prices for a week, then add a line chart to show the price movement.
```

### Test 8: Pie Chart
```
Create an Excel file named "market_share.xlsx" with company market share data, then add a pie chart to visualize the distribution.
```

## CSV Integration Tests

### Test 9: CSV to Excel Conversion
```
Convert this CSV data to Excel:
Name,Age,City
Alice,28,New York
Bob,32,Los Angeles
Charlie,25,Chicago

Save it as "people_data.xlsx"
```

### Test 10: Excel to CSV Export
```
Create an Excel file with sample data, then export it to CSV format.
```

## Complex Business Scenarios

### Test 11: Sales Report with Charts
```
Create a comprehensive sales report Excel file named "monthly_report.xlsx" with:
- Sheet 1: Sales data with columns [Month, Product, Units Sold, Revenue]
- Sheet 2: Summary with totals
- Add a bar chart showing revenue by month
- Apply professional formatting
```

### Test 12: Employee Directory
```
Create an employee directory Excel file with:
- Columns: Name, Department, Email, Phone, Hire Date
- Sample data for 5 employees
- Professional formatting with headers
- Auto-filter enabled
```

### Test 13: Budget Tracker
```
Create a budget tracker Excel file with:
- Income and expense categories
- Monthly data for 3 months
- Summary calculations
- Conditional formatting for over-budget items
```

## Error Handling Tests

### Test 14: Invalid Filename
```
Create an Excel file with a very long filename that exceeds normal limits.
```

### Test 15: Empty Data
```
Create an Excel file with no headers and no data to test error handling.
```

## Performance Tests

### Test 16: Large Dataset
```
Create an Excel file with 1000 rows of sample data to test performance.
```

### Test 17: Multiple Sheets
```
Create an Excel file with 5 different sheets, each containing different types of data.
```

## Integration Tests

### Test 18: Download Link Verification
```
Create any Excel file and verify that the download link works correctly in your browser.
```

### Test 19: Session Persistence
```
Create multiple Excel files in sequence to test if session persistence works between requests.
```

### Test 20: Format + Chart Combination
```
Create an Excel file with custom formatting AND charts to test if both features work together.
```

## Expected Results Checklist

For each test, verify:
- ✅ Excel file is created successfully
- ✅ Download link is provided and works
- ✅ File contains correct data
- ✅ Formatting is applied (if requested)
- ✅ Charts are created (if requested)
- ✅ No error messages in response
- ✅ File can be opened in Excel/Google Sheets

## Troubleshooting Tips

If a test fails:
1. Check the MCP server logs: `docker compose logs exel-mcp --tail=20`
2. Verify file exists: `ls -la /home/ghost/bin/docker/exel_mcp/output/`
3. Test download link manually in browser
4. Check session initialization in logs
5. Verify Open-WebUI tool configuration is up to date

## Advanced Test Scenarios

### Test 21: Real-world CV Processing
```
Take a real CV text and convert it to a structured Excel file with sections for Education, Experience, Skills, and Certifications.
```

### Test 22: Data Analysis Report
```
Create sample sales data and generate an Excel report with summary statistics, charts, and insights.
```

### Test 23: Multi-language Support
```
Create an Excel file with non-English characters and special symbols to test encoding support.
```

Use these prompts systematically to validate all aspects of your Excel MCP integration!