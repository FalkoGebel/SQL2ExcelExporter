# SQL 2 Excel Exporter
A Program to export data from SQL tables to Excel files.

## Basic structure
- Library project for the UI-independent functionalities
- Library test project for the test for the library project (MSTest with fluent assertion)
- UI project for presentation - using WPF

## Programming Language
- C#

## Status
- Choose SQL server, SQL database and SQL table
- Select the columns to export (unsupported data types are recognizable)
- Choose the directory to create the Excel file in
- Define the styles for the header line and the data lines
- Hint: for large tables with many lines the file creation process needs a lot of time - at the moment there is no progress information

![SQL 2 Excel Exporter UI](/README-Images/UI.jpg?raw=true "SQL 2 Excel Exporter")

![Information Box](/README-Images/Information.jpg?raw=true "SQL 2 Excel Exporter")

![Result Libre Calc](/README-Images/Result_Libre_Calc.jpg?raw=true "SQL 2 Excel Exporter")