# Dashboard Project

This project is a data analysis and visualization dashboard built using Microsoft Excel and VBA macros. The dashboard facilitates data import, cleaning, transformation, and visualization, providing an efficient way to gain insights from data. It is structured to handle raw data and generate meaningful visual representations and summaries.

## Features

1. Data Import:
- Automatically imports raw data for analysis.
- Supports various formats compatible with Excel.

2. Data Cleaning:
- Ensures data consistency by removing duplicates and handling missing values.
- Transforms raw data into a standardized format.

3. Data Transformation:
- Processes cleaned data to generate meaningful insights.
- Creates structured datasets ready for analysis.

4. Pivot Tables:
- Provides summarized views of the data.
- Includes two pre-built pivot tables for key metrics and trends.

5. Visualizations:
- Generates clear and interactive charts for data exploration.
- Includes two pre-built visualizations for highlighting critical insights.

6. Automation with VBA:
- Streamlines the workflow using custom VBA macros.
- Reduces manual effort and minimizes errors.

## File Structure
The project consists of the following sheets:
- Import: Sheet for importing raw data.
- Raw Data: Contains the unprocessed data as imported.
- Cleaned Data: Data after cleaning and preprocessing.
- Transformed Data: Data prepared for visualization and pivot table creation.
- Pivot Table 1 and Pivot Table 2: Summary tables for key data metrics.
- Visualisasi 1 and Visualisasi 2: Graphs and charts to visualize insights.

## Key Results
1. Cleaned Data Output:
- The dataset is cleaned by removing duplicates based on the Invoice ID column.
- Rows with empty or invalid values are deleted.
- The date column is reformatted to a consistent dd-mm-yyyy format.

2. Data Transformation:
- Additional columns such as Net Income, Sales Category, and Transaction Time Category are added for advanced analysis.
- Net Income is calculated as the difference between Total Sales and COGS.

3. Pivot Tables:
- Pivot Table 1: Summarizes total sales based on Branch and Product Line.
- Pivot Table 2: Displays total sales and profit categorized by Transaction Time Category and Branch.

4. Visualizations:
- Sales by Branch and Product Line Chart: Provides a comparison of total sales across branches and product lines.
- Sales by Time and Branch Chart: Shows sales trends based on transaction times.

## Prerequisites
- Microsoft Excel (2016 or later).
- Basic understanding of VBA (optional, for customization).

## How to Use
- Open the Excel file.
- Use the Import sheet to load your raw data.
- Run the provided macros to clean and transform the data.
- Explore the pivot tables and visualizations for insights.

