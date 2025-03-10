'Add a reference to the Microsoft Office Interop Excel library:
'Right-click on your project in the Solution Explorer And select "Add" > "Reference..."
'Microsoft Excel Object library

'Project -> Add Componentent -> Application Configuration File
'Application Level Settings are stored in a configuration file. Project + Add Reference, select "System.Configuration".
' For executable programs this file Is located in the same directory as the .exe And Is named after the assembly, Or executable.
' Example: MyAssembly.config, Another.Assembly.config
'User Scoped Settings are stored in the user profile usr.config
'Example: C:\Documents and Settings\USERNAME\Local Settings\Application Data\ApplicationName
'Connection Strings are stored in the <connectionStrings></connectionStrings> section in your config file.
'<connectionStrings>
'<clear />
'<add name = "Name"
'providerName = "System.Data.ProviderName"
'connectionString = "Valid Connection String;" />
'</connectionStrings>
'sqlComm.Connection.ConnectionString = ConfigurationManager.ConnectionStrings["MyWebService.Properties.Settings.dbConnString"].ConnectionString;


Imports System.Configuration
Imports System.Data.SqlTypes
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim occupation As String
        occupation = ConfigurationManager.AppSettings("occupation")
        'ConfigurationManager.AppSettings["MySetting"] = "SomeStuff";
        'var conn = ConfigurationManager.ConnectionStrings["DevSqlServer"];


        ' Get the SQLquery file:
        'Dim sqlAllStrings As New System.IO.StreamReader("SQLqueries.txt")
        'Dim sqlAllStrings As New System.IO.File.ReadAllText("SQLqueries.txt")
        ''Dim subjectsWithLines() As String = Split(sqlAllStrings, Chr(10))
        'Dim lines As String() = sqlAllStrings.Split(New String() {vbCrLf + vbCrLf}, StringSplitOptions.None)
        ' ATTENTION: empty line should not have a space.
        ''First line of query is number and description "31 Text" 
        Dim SqlString = Right(Str, Len(Str) - InStr(Str, vbCrLf) + 1)
        Dim SqlHeader = Left(Str, InStr(Str, vbCrLf) - 1)
        ''Dim splittedStr As String() = Split(sqlString, vbCrLf)
        ''slqString = String.Join(vbCrLf, splittedStr, 1, splittedString.Length)


        Label1.Text = occupation

        ' Create a new instance of Excel application
        Dim excelApp As New Excel.Application

        Dim workbook As Excel.Workbook
        Dim worksheet As Excel.Worksheet

        ' Make the Excel application visible
        excelApp.Visible = True

        ' Add a new workbook
        workbook = excelApp.Workbooks.Add()

        ' Access the first worksheet
        worksheet = CType(workbook.Sheets(1), Excel.Worksheet)
        worksheet.Name = "MySheet"

        ' Write some data to the worksheet
        worksheet.Cells(1, 1).Value = "Hello, Excel!"

        'xlApp = New Excel.ApplicationClass
        'xlWorkBook = xlApp.Workbooks.Open("c:\test1.xlsx")
        'xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        'display the cells value B2
        'MsgBox(xlWorkSheet.Cells(2, 2).value)
        'edit the cell with new value
        'xlWorkSheet.Cells(2, 2) = "http://vb.net-informations.com"
        'xlWorkBook.Close()
        'xlApp.Quit()

        'releaseObject(xlApp)
        'releaseObject(xlWorkBook)
        'releaseObject(xlWorkSheet)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
