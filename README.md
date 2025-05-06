# Resultviewer
## Description
### SQL Dialect


#### DAO/ADO/Jet
* DAO (Data Access Objects): ACE/Jet's natural interface layer. Deprecated in favor of ADO.
* ADO (ActiveX Data Objects): Uses OLE DB (Object Linking and Embedding) programming interface. OLE DB is based on the Microsoft Component Object Model (COM).
* The resultviewer uses currently ADO for MS Access databases (.mdb) and Excel

#### Wildecard characters in LIKE: * or %

* Depends on the "ANSI Query Mode" of the used interface. These modes are specific to ACE/Jet (but ressemble ANSI/ISO SQL-89 and SQL-92 Standards).
* DAO interface uses ANSI-89 Query Mode ('traditional mode'): * 
* ADO interface (OLE DB) uses ANSI-92 Query Mode ('SQL Server compatibility mode'): %
* In ODBC the query mode can be explicitly specified via the ExtendedAnsiSQL flag
* The MS Access user interface, from the 2003 version onwards, can use either query mode. So be aware: A query may work in MS Access unserinterface, but not with DAO or ADO
* ACE/Jet SQL syntax has an ALIKE keyword, which allows the ANSI-92 Query Mode characters (% and _) regardless of the query mode of the interface
* 