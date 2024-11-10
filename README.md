# SalesDataAutomation - RPA Excel Automation Project

## Overview

SalesDataAutomation is a robotic process automation (RPA) project developed using UiPath to automate data processing tasks involving Excel files. The project aims to streamline data collection from multiple sources, merge the information, and perform advanced data analysis, including pivot tables and chart generation. This repository demonstrates how RPA can simplify tedious data operations, making them faster, more accurate, and easier to manage.

## Features
- Downloads three Excel files dynamically from Google Drive using HTTP requests and stores them in a dynamically created folder named **SalesDataAutomation**.
- Merges data from multiple Excel files into a consolidated **SalesMIS.xlsx** file.
- Performs data transformation by converting the consolidated data into an Excel table.
- Generates a pivot table from the merged data to summarize key information.
- Creates a pie chart for visual representation of the processed data.
- Implements modular workflows for better project structure and easy maintenance.

## Folder Structure

- **SalesDataAutomation/**: This folder is created dynamically to store the downloaded Excel files and the final **SalesMIS.xlsx** file.
- **Workflows**: The project is divided into multiple workflows for easier readability and debugging. The workflows used include:
  - **Folder.xaml**: Creates the **SalesDataAutomation** folder.
  - **Excel File Download.xaml**: Downloads the three Excel files (**Furniture Outlet 1**, **Furniture Outlet 2**, and **Furniture Outlet 3**) from Google Drive.
  - **SalesMIS.xaml**: Creates the consolidated **SalesMIS.xlsx** file in the **SalesDataAutomation** folder.
  - **Appending 2.0.xaml**: Appends the data from the downloaded Excel files into **SalesMIS.xlsx**.
  - **Clear Sheet.xaml**: Clears existing data from **SalesMIS.xlsx** before appending new data (if applicable).
  - **Formatting.xaml**: Converts the appended data into an Excel table.
  - **Pivot.xaml**: Creates a pivot table based on the formatted data.
  - **Pie Chart.xaml**: Generates a pie chart from the pivot table.

## Flowchart Overview

The main automation process is structured in a sequence, which ensures stepwise execution of the entire automation:
1. **Folder Creation**: Using **Folder.xaml**, a folder named **SalesDataAutomation** is created.
2. **Parallel Processing**: Multiple workflows are run in parallel to enhance efficiency.
   - **Excel File Download.xaml** downloads the input files.
   - **SalesMIS.xaml** creates the empty **SalesMIS.xlsx** file.
3. **Excel Automation**:
   - Data from the downloaded Excel files is read and appended into **SalesMIS.xlsx**.
   - The data is then formatted into a table, a pivot table is created, and a pie chart is generated to summarize the data.

## How It Works
1. **Downloading Files**: The project starts by dynamically downloading three Excel files - **Furniture Outlet 1**, **Furniture Outlet 2**, and **Furniture Outlet 3** - using direct download links from Google Drive. The files are stored in a dynamically created folder called **SalesDataAutomation**.
2. **Data Consolidation**: A new Excel file named **SalesMIS.xlsx** is created in the same folder, and the data from the downloaded files is appended to it.
3. **Data Transformation**: The data in **SalesMIS.xlsx** is transformed into a tabular format for easier analysis.
4. **Pivot Table and Chart**: A pivot table is created to summarize the data, and a pie chart is generated for visual representation.

## Project Files
The repository includes the following files:
- **Main.xaml**: The main workflow that orchestrates the entire automation process.
- **Folder.xaml**: Workflow to create the necessary folder.
- **Excel File Download.xaml**: Workflow to download the files.
- **SalesMIS.xaml**: Workflow to create the **SalesMIS.xlsx** file.
- **Appending 2.0.xaml**: Workflow to append data.
- **Clear Sheet.xaml**: Workflow to clear existing data.
- **Formatting.xaml**: Workflow to format the appended data.
- **Pivot.xaml**: Workflow to create the pivot table.
- **Pie Chart.xaml**: Workflow to generate the pie chart.
- **SalesDataAutomation** folder: Stores the input Excel files and **SalesMIS.xlsx**.

## Prerequisites
- UiPath Studio with Modern Design Experience enabled.
- Microsoft Excel installed on the system.
- Direct download links for the Excel files stored in Google Drive.

## Setup Instructions
1. Clone the repository to your local machine.
2. Open the project in UiPath Studio.
3. Update the download links in the **Excel File Download.xaml** workflow to point to your own Google Drive files if necessary.
4. Run the **Main.xaml** workflow to execute the automation.

## License
This project is licensed under the MIT License - see the LICENSE file for details.

