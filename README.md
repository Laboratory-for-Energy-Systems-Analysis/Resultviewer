# Resultviewer (SQL in Excel)

> Leverage the power of SQL in Excel for data extraction and charting.

> Special features for multi-regional energy system models with output of different scenarios over time steps

## Purpose

> The Resultviewer is an MS Excel file with macros inside that executes an arbitrary number of SQL-queries on tabular input data, and puts the extracted result tables together with an optional chart on a new Excel-sheet.
- Currently implemented input tables: A  sheet in another Excel-file (.xslx), or a table of an MS Access database (.mdb)
- The SQL-queries are specified in a separate text file
- The Resultviewer is optimized for speed. It can put houndreds of charts in seconds.
- Using VBA-macros, you can easily customize the Resultviewer to your needs, e.g. by changing the layout of the charts


## Capabilities
- **Chart customization:** Type (column, line etc.), color, order of rows, labels, size, second axis
- **Multiple charts by scenario and region:** The same SQL-query can be executed for several scenarios and regions ('scenario' and 'region' are field-header-column in the input table). For this, you write in the SQL-query `<SCENARIO>` and `<REGION>`, and this templates will substituted automatically.
- **Templates for time steps:** To make the writing of SQL-queris with many time steps as table-columns more simple, a template can be used in the SQL-queries, such that only a single time-step (template) has to be entered manually.
- **Comparison of scenarios:** In the Resultviewer, a new scenario is put on a new sheet. For better comparison, vertical axes can be levelized across sheets, and charts of three scenarios can be put on a single sheet for side-by-side comparison.

## Description/How to use



To control the import, the Resultviewer employs a panel and two Excel sheets:
1. A **Control panel**. To access: Go to the ribbon of Excel --> Add-in. You see  a row of buttons. Press `Panel`, which opens the control panel of the Resultviewer
2. A **setting** sheet, and a **color and order** sheet. The actual names of the sheets in Excel can be determined in the panel, such that different settings can be used; for our global energy system model, we use for example the names `SETTING` and `COLOR&ORDER`, and for the example the names are chosen to be `SETTING_Example` and `COLOR&ORDER_Example`.  

### Run the example

1. Download from this site the 3 files: 
	- `Resultviewer_NT.xls`
	- `sql_queries_Example.txt`
	- `WEO2024.xlsx`
2. Open `Resultviewer_NT.xls`, and enable macros (if not already done)
3. In the Resultviewer, go to sheet `COLOR&ORDER_Example`. Look for the table a little bit on the right that has the paths to `WEO2024.xlsx` and `sql_queries_Example.txt`. Adjust the paths to the locations of your files. 
5. In the control panel, adjust the name of the color & ordering sheet to `COLOR&ORDER_Example`, and press `Load` to load values from the aforementioned table into the panel. Now, the panel should have the correct values of the example. Clearly, you could have entered the values also manually directly in the control panel.
6. In the control panel, select a scenario in the drop-down box
7. In the control panel, press `View Scenario`


## Colors

Colors can be fully customized. In the Resultviewer, colors are specified by choosing a number between 1-57 of Excel's palette (i.e., the colorindex in Excel). The following options exist:

- **Standard color scheme:** Colors for each chart series can be specified in the `COLOR&ORDER` sheet in column `D`. To see the color of a colorindex that you want to choose, the table `Excel Palette` lists all colors. Press `Update Colors` in the ribbon to see all colors updated.
- **Alternative color scheme:** An alternative scheme can be entered in the next column `E` on the `COLOR&ORDER` sheet. In the panel, you can tick whether this alternative colors should be used for chart import. As a special feature of this alternative scheme, in next column `F` you can enter a letter `s` such that the color is shaded with diagonal strips. 
- **Custom Excel color palette:** If you need other colors than available in the standard Excel palette, you can specify your own palette in table `Palette` on the `COLOR&ORDER` sheet by re-defining a colorindex 1-57 to a RGB value (e.g. `255,0,0` for red). Press `use palette (Color&Order sheet)` in the panel to use the new definitions. Note: Not all colorindices have to be re-defined, and any order of indices is allowed.
- **Coloring the interior of a selected Excel object:** If you need to alter a color of a single object (e.g. a chart series in a single chart), press `Set color` in the ribbon. This will open a dialog, where you can select the colorindex. 


## Issues (ToDo)
 - The start of SQL queries in the text file is identified by the SQL keywords `TRANSFORM` or `SELECT`. Currently the code searches for uppercase keywords only (SQL syntax is case-insensitive).

## FAQ

- **Why use Excel?** You have both chart and data available, and you have all the advanced Excel features available for post-calculations and copy-pasting.
- **Why is the Resutviewer in .xls format, and not in .xslm format?** Indeed, you can save the .xls as .xlsm  without any problem. We found that the .xls format is more stable if many charts are generated and also for VBA macros, despite that .xlsm is a newer format, and should be more stable as advertised. Also .xls is faster, perhaps also due to the fact that .xls is limited to 65k rows, whereas .xlsm can have 1M rows.  
- **Why use macros (VBA, Visual Basic for Application), instead of a newer connection interface to Excel, by using Python/VB/C# etc.?** By our tests, VBA is up to three times faster then the other connections, which is due to the so-called Marahalling (VSTO). Also, the VBA code is directly accesible inside Excel, and can be adapted by the user.


## Limits
- Names in the column-header "scenario" should be less than 26 characters long because they will become sheet names (32 char. limit): If dubplicated, a number in parentheses will be appended, e.g. (2)
- Not more than approx. 1000 charts in Excel. With more, Excel seems to break-down (out of memory, or, file cannot be saved etc.). 1000 charts is easily reached, e.g. with 20 regions, 20 scenarios, 30 charts per scenario. 
- If you have several Resultsviewers open at the same time, Excel may become unstable, and files cannot be saved.


### JET-engine limits

For queries on Excel-sheets and MS Access, the Microsoft ACE/JET-engine with interface AOD is used. This engine has limits, but has also the advantage of the PIVOT-table functionality by using the keyword `TRANSFORM` as is used in the example. The limits are: 
 
- CASE and nested IFF statements can have maximal 15 levels (otherwise the SQL-engine gives an error: "SQL query too complex").
- Comments inside SQL queries are not allowed
- A query cannot conists of several sequential SELECT statements, separated by a semicolon, as would be possible in other DB engines
- The SQL syntax and table headers ar case-insensitive
- The used interface AOD uses % in LIKE (and not * as is possible MS Access)


## Background: DAO/ADO/Jet etc.
- The resultviewer uses currently ADO for MS Access databases (.mdb) and Excel.
- ADO (ActiveX Data Objects): Uses OLE DB (Object Linking and Embedding) programming interface. OLE DB is based on the Microsoft Component Object Model (COM). The older DAO (Data Access Objects) is deprecated in favor of ADO.
- ACE/Jet has an "ANSI Query Mode" (ressembles ANSI/ISO SQL-89 and SQL-92 Standards). ADO interface (OLE DB) uses ANSI-92 Query Mode ('SQL Server compatibility mode'). The MS Access user interface, from the 2003 version onwards, can use either query mode (* or % in LIKE). So be aware: A query may work in MS Access unserinterface, but not with DAO or ADO In ODBC the query mode can be explicitly specified via the ExtendedAnsiSQL flag. ACE/Jet SQL syntax has an ALIKE keyword, which allows the ANSI-92 Query Mode characters (% and _) regardless of the query mode of the interface

