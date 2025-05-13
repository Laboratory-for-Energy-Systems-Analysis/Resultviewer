# Resultviewer

## Purpose

* The Resultviewer is an Excel-file with macro-code behind that executes SQL-queries on tabular input data (another .xslx Excel-sheet or an Access-DB) and puts the resulting tables with a charts on an Excel-sheet.
The charts can be customized with a template: Type (column, line etc.), colors, ordering, labeling.
* The Resultviewer is geared towards executing the same SQL-query for different scenarios and for different regions (where 'scenario' and 'region' are field-column in the input table).
* The Resultviewer is geared also to show data over time by providing a template in the SQL-queries, such that not all time-steps have to be specified manually.

## Description


## FAQ

* **Why is the Excel provides in .xls format, and not in the newer .xslm?** You can save the .xls as .xlsm, without any problem. We ahve found that the .xls format is more 
          stable with many charts and for VBA, despite taht .xlsm is a newer format, and should be more stable as advertised. Also .xls is much faster, perhaps also due to the fact, that .xls are limited to 65k rows, and .xlsm to 1M rows.  
* **Why use Excel?** You have both chart and data available, and you have all the advanced Excel features available for post-calculations and copy-pasting. 
* **Why use the VBA (Visual Basic for Application), instead of a newer connection types, for example Python/VB/C# <--> Excel?:** By our tests, VBA is up to three times faster then the other connections, which is due to Marahalling overhad. Also, the VBA code is directly accesible inside Excel. 
* **How many charts can you put into Excel?** By our experience, after approx. 1000 charts, Excel seems to break-down (out of memory, or, file cannot be saved etc.). This is easily reached, e.g. with 20 regions, 20 scenarios, 30 charts per scenario. 
* **Can you have several Resultsviewers open?** Yes, but mind the limit of the number of charts. Excel may go out of memory, and files cannot be saved.

## Caveats
* Names in the column scenario should be less than 26 characters long, because they become sheet names, occasionally with a number at the end (XX)


### SQL Dialect


#### DAO/ADO/Jet
* The resultviewer uses currently ADO for MS Access databases (.mdb) and Excel
* ADO (ActiveX Data Objects): Uses OLE DB (Object Linking and Embedding) programming interface. OLE DB is based on the Microsoft Component Object Model (COM).
* DAO (Data Access Objects): ACE/Jet's natural interface layer. Deprecated in favor of ADO.

#### Wildecard characters in LIKE: * or %

* Depends on the "ANSI Query Mode" of the used interface. These modes are specific to ACE/Jet (but ressemble ANSI/ISO SQL-89 and SQL-92 Standards).
* DAO interface uses ANSI-89 Query Mode ('traditional mode'): * 
* ADO interface (OLE DB) uses ANSI-92 Query Mode ('SQL Server compatibility mode'): %
* In ODBC the query mode can be explicitly specified via the ExtendedAnsiSQL flag
* The MS Access user interface, from the 2003 version onwards, can use either query mode. So be aware: A query may work in MS Access unserinterface, but not with DAO or ADO
* ACE/Jet SQL syntax has an ALIKE keyword, which allows the ANSI-92 Query Mode characters (% and _) regardless of the query mode of the interface

#### Limitation of the JET-engine
* Using the JET-engine, the SELECT and nested IFF statements can have maximal 15 levels (otherwise the SQL-engine gives an error: "SQL query too complex").
* Comments inside SQL queries are not allowed
* A query cannot conists of several sequential SELECT statements, separated by a semicolon, as in other DB engines