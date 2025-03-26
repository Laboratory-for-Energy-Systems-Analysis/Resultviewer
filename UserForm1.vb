'Add a reference to the Microsoft Office Interop Excel library:
'Right-click on your project in the Solution Explorer And select "Add" > "Reference..."
'Microsoft Excel Object library

'Add a reference to ADODB:
'Right-click on your project in the Solution Explorer And select "Add" > "Reference..."
'Microsoft ActiveX Data Objects 6.1 Library

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



'The workaround here Is easy:
'Just open the AppDesigner For the VB-Project Properties And change the Startup Form To a valid setting. Recompile, done. This needs only To be done, When the Start Form gets renamed.

Imports System.Configuration
Imports System.Runtime.InteropServices
Imports ADODB
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel





Public Class UserForm1

    Dim regionArray() As String
    Dim allRegions As String

    Public con As ADODB.Connection 'early binding
    Public rs As ADODB.Recordset 'early binding

    Public excelApp As Excel.Application
    Public excelWorkBook As Excel.Workbook

    'Public excelWorkSheets As Excel.Worksheets
    'Public excelWorkSheetOne As Excel.Worksheet
    'Public excelActiveSheet As Excel.Worksheet


    Sub OpenExcelFile_Click(sender As Object, e As EventArgs) Handles OpenExcelFile.Click

        'Dim fn As String

        'excelApp = GetObject(, "Excel.Application")
        excelApp = New Excel.Application
        excelApp.Visible = True
        excelWorkBook = excelApp.Workbooks.Open(ResultviewerFileNameBox.Text)
        'excelWorkSheets = excelWorkBook.Worksheets()
        'excelWorkSheetOne = excelWorkBook.Worksheets("Sheet1")
        'excelActiveSheet = excelApp.ActiveSheet
        'display the cells value B2
        MsgBox(excelWorkBook.Sheets("Sheet1").Cells(2, 2).value)


        'If UserForm1 IsNot Nothing Then MsgBox("UserForm1 is not initialized.")
        'End If
        If excelWorkBook Is Nothing Then MsgBox("excelWorkBook is not initialized.")
        If excelWorkBook IsNot Nothing Then MsgBox("excelWorkBook is initialized.")
        If excelApp.ActiveChart IsNot Nothing Then MsgBox("excelApp.ActiveChart is not initialized.")




        'edit the cell with new value
        'WorkSheet.Cells(2, 2) = "test"
        'WorkSheet.Close()
        'excelApp.Quit()



    End Sub

    Private Sub CreateExcelFile_Click(sender As Object, e As EventArgs) Handles CreateExcelFile.Click

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
        'Dim SqlString = Right(Str, Len(Str) - InStr(Str, vbCrLf) + 1)
        'Dim SqlHeader = Left(Str, InStr(Str, vbCrLf) - 1)
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




    Private Sub revArrScnButton_Click() Handles revArrScnButton.Click

        'Copy charts from different scenarios back from JAZZ

        Call ReverseArrChartsScen()


    End Sub


    Private Sub arrScenButton_Click() Handles arrScenButton.Click

        'Copy charts of different scenario to new sheet

        Call ArrChartsScen()

    End Sub

    Private Sub copyDBAndAggregateButton_Click() Handles CopyDBandAggregateButton.Click

        Call CopyDBandAggregate(FilenameBox.Text)

    End Sub

    Private Sub arrangeButton_Click() Handles arrangeButton.Click

        Call ArrangeMyCharts()

    End Sub

    Private Sub secondaryButton_Click() Handles secondaryButton.Click

        Call SecondaryAxis()

    End Sub

    Private Sub equalizeButton_Click() Handles equalizeButton.Click

        Call EqualizeAxes()

    End Sub

    Private Sub ConvertToBWButton_Click() Handles ConvertToBWButton.Click

        Call BW(excelApp)

    End Sub

    Private Sub selectAllRegion_Click() Handles selectAllRegion.Click

        Dim i As Integer
        Dim c As Excel.Range

        'Iterate to search for chart numbers in column C, and then put "*" (if not "NA") in column B
        For i = 1 To 100
            c = excelWorkBook.Sheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then  'CHANGE
                If c.Offset(0, -1).Value <> "NA" Then
                    c.Offset(0, -1).Value = "*"
                End If
            End If
        Next i

    End Sub

    Sub selectAll_Click() Handles selectAll.Click

        Dim i As Integer
        Dim c As Excel.Range

        'Iterate to search for chart numbers in column C, and then put "*" (if not "NA") in column A
        For i = 1 To 1000
            c = excelWorkBook.Sheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then 'CHANGE
                If c.Offset(0, -2).Value <> "NA" Then
                    c.Offset(0, -2).Value = "*"
                End If
            End If
        Next i

    End Sub

    Sub DeselectAll_Click() Handles DeselectAll.Click

        Dim i As Integer
        Dim c As Excel.Range

        For i = 1 To 100
            c = excelWorkBook.Sheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                If c.Offset(0, -2).Value <> "NA" Then
                    c.Offset(0, -2).Value = ""
                End If
                If c.Offset(0, -1).Value <> "NA" Then
                    c.Offset(0, -1).Value = ""
                End If
            End If
        Next i

    End Sub

    Sub unloadButton_Click() Handles unloadButton.Click

        Close()

    End Sub

    Sub colorButton_Click() Handles colorButton.Click

        Dim i As Integer
        Dim val
        Dim val2
        Dim valChck
        'Dim colArray As Integer()
        Dim wf = excelApp.WorksheetFunction


        For i = 3 To 1000
            val = excelWorkBook.Sheets("Sheet1").Cells(i, 5).Value
            val2 = excelWorkBook.Sheets("Sheet1").Cells(i, 6).Value
            valChck = excelWorkBook.Sheets("Sheet1").Cells(i, 3).Value
            'If i < 20 Then
            '   MsgBox val2
            'End If
            If Not wf.IsNumber(valChck) Then
                If RGBCheckBox.Checked Then
                    If wf.IsNumber(val2) Then
                        If (0 < val2 And val2 < 58) Then
                            excelWorkBook.Sheets("Sheet1").Cells(i, 6).Interior.ColorIndex = CInt(val2)
                        Else
                            MsgBox("ColorIndex " & val2 & " of custom palette is out of range.")
                        End If
                    ElseIf Not String.IsNullOrEmpty(val2) Then
                        MsgBox("ColorIndex " & val2 & " is not a single, integer number.")
                    End If
                ElseIf wf.IsNumber(val) Then
                    If (0 < val And val < 58) Then
                        excelWorkBook.Sheets("Sheet1").Cells(i, 5).Interior.ColorIndex = CInt(val)
                    Else
                        MsgBox("ColorIndex " & val & " of default palette is out of range.")
                    End If
                ElseIf Not String.IsNullOrEmpty(val) Then
                    MsgBox("ColorIndex in row " & i & "is not empty but not a single, integer number.")
                    MsgBox("ColorIndex " & val & " is not a single, integer number.")
                End If
            End If
        Next i

    End Sub

    Sub avgButton_Click() Handles AvgButton.Click

        Call avgBaseYear()

    End Sub

    Sub ResetColorButton_Click() Handles ResetColorButton.Click
        'Workbook.ResetColors()
        Call ResetColorPalette(excelWorkBook)
    End Sub

    Private Sub UserForm_Click()

    End Sub

    Sub WECcolorButton_Click() Handles WECcolorButton.Click
        Call setColorPalette()
        'Call WECcolorPalette
    End Sub

    Sub FilenameButton_Click() Handles FilenameButton.Click

        Dim fileToOpen As String

        '    ChDrive "M:\"
        '    ChDir "M:\GMM\Phase IV (WEC 2019)\Answer Database"

        'fileToOpen = Application.GetOpenFilename("Answer database files (*.mdb),*.mdb", , "Open a database...")

        Dim fd As New OpenFileDialog()

        'fd.Title = "Open File Dialog"
        'fd.InitialDirectory = "C:\"
        'fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        'fd.FilterIndex = 2
        'fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            fileToOpen = fd.FileName
            FilenameBox.Text = CStr(fileToOpen)
            Call updateCases(answerFileName:=CStr(fileToOpen))
        Else
            MsgBox("VBA-Application could not open the file.")
        End If
    End Sub

    Sub updateCases(answerFileName As String)

        ' Get information from Access-DB on the available ANSWER-cases

        Dim cnt As New ADODB.Connection
        Dim rst As New ADODB.Recordset
        Dim rst2 As New ADODB.Recordset
        Dim i As Integer
        Dim regionStr As String
        Dim fieldNames() As String
        'Dim stepTmp As Integer
        'Dim step As Integer

        ' Open connection to the database
        cnt.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & answerFileName & ";")

        ''When using the Access 2007 Northwind database
        ''comment the previous code and uncomment the following code.
        'cnt.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        '    "Data Source=" & strDB & ";"

        ' Open recordset based on Orders table
        rst.Open("SELECT DISTINCT CaseName FROM tblTRESULTS", cnt)

        CaseNameComboBox.Items.Clear()
        If rst.BOF And rst.EOF Then
            MsgBox("No scenarios in selected database.")
        Else
            rst.MoveFirst()
        End If
        Do While Not rst.EOF
            CaseNameComboBox.Items.Add(rst.Fields(0).Value)
            rst.MoveNext()
        Loop
        CaseNameComboBox.SelectedIndex = 0

        rst.Close()
        rst.Open("SELECT DISTINCT Region FROM tblTRESULTS", cnt)
        If rst.BOF And rst.EOF Then
            MsgBox("No regions.")
        End If
        i = 0
        Do While Not rst.EOF
            'Debug.Print i
            regionStr = rst.Fields(0).Value
            If ((StrComp("REGION0", regionStr) <> 0) And (StrComp("_GLOBAL", regionStr) <> 0)) Then
                i = i + 1
                ReDim Preserve regionArray(0 To i - 1)  'CHANGE previous: 1 TO i
                regionArray(i - 1) = regionStr     'CHANGE: previous i 
            End If
            rst.MoveNext()
        Loop

        rst.Close()
        rst.Open("SELECT * FROM tblTRESULTS WHERE Parameter = 'U.TOTCOST'", cnt)
        ReDim fieldNames(0 To (rst.Fields.Count - 1))
        For i = 0 To rst.Fields.Count - 1
            fieldNames(i) = rst.Fields(i).Name
        Next
        eop = CInt(Mid(fieldNames(rst.Fields.Count - 1), 2))
        i = 0
        For Each s In fieldNames
            If StrComp("T2", Mid(s, 1, 2)) = 0 Then
                If i = 0 Then
                    sop = CInt(Mid(s, 2))
                End If
                If i = 1 Then
                    timestep = CInt(Mid(s, 2)) - sop
                End If
                i = i + 1
            End If
        Next
        MsgBox("Start year = " & CStr(sop) & ", Step = " & CStr(timestep) & ", End year = " & CStr(eop))
        TimeHorizonDBbox.Text = CStr(eop)
        rst.Close()
        cnt.Close()
        rst = Nothing
        cnt = Nothing

        i = 1
        For Each s In regionArray
            If i = 1 Then
                allRegions = "'" & s & "'"
            Else
                allRegions = allRegions & ", '" & s & "'"
            End If
            i = i + 1
        Next s
        MsgBox("Regions = " & allRegions)

    End Sub

    Public Sub OnEnd()

        excelApp.ScreenUpdating = True
        excelApp.EnableEvents = True
        excelApp.DisplayStatusBar = True
        excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

    End Sub

    Public Sub OnStart()

        excelApp.ScreenUpdating = False
        excelApp.EnableEvents = False
        excelApp.DisplayStatusBar = False
        excelApp.Calculation = Excel.XlCalculation.xlCalculationManual

    End Sub

    Sub ImportAllButton_Click() Handles importAllButton.Click

        ' Batch import of several cases
        ' The names of the cases are in a list with text "Batch" in the header in "Sheet1"

        Dim caseList As Excel.Range
        Dim c As Excel.Range
        Dim longName As String

        c = excelWorkBook.Sheets("Sheet1").Range("A1:Z100").Find("Batch", LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)
        If c Is Nothing Then
            MsgBox("Could not find list of cases in Range A1:Z100. The list as 'Batch' as header.")
        Else
            caseList = excelWorkBook.Sheets("Sheet1").Range(c.Offset(1, 0), c.End(XlDirection.xlDown))
            If Not selectRegions.Checked Then
                Call OnStart()
            End If
            For Each r In caseList.Rows
                'Debug.Print caseName & " --- " & caseNameLong & " --- "
                longName = r.Cells(1, 2).Value 'subtitle in charts
                If Not longName = "" Then
                    longName = "(" & longName & ")"
                End If
                Call ImportFromMDB(answerFileName:=FilenameBox.Text,
                        caseName:=r.Cells(1, 1).Value,
                        scenarioName:=r.Cells(1, 3).Value,
                        caseNameLong:=longName)
            Next r


            MsgBox("Generation & Arranging successful.")

            If Not selectRegions.Checked Then
                Call OnEnd()
            End If

        End If

    End Sub

    Sub ImportButton_Click() Handles importButton.Click

        If Not selectRegions.Checked Then
            Call OnStart()
        End If

        Call ImportFromMDB(answerFileName:=FilenameBox.Text,
                        caseName:=CaseNameComboBox.Text,
                        scenarioName:=ScenarioTextBox.Text)

        MsgBox("Import & Charting successful.")

        If Not selectRegions.Checked Then
            Call OnEnd()
        End If

    End Sub



    Sub ImportFromMDB(answerFileName As String,
                    caseName As String,
                    scenarioName As String,
                    Optional caseNameLong As String = "")

        Dim oSheet As Excel.Worksheet
        Dim i As Integer
        Dim oRange As Excel.Range
        Dim sheetName As String
        Dim sheetName2 As String
        'Dim nameExists As Boolean
        'Dim countries As String
        'Dim categories As String
        Dim str As String
        Dim lChartType As Excel.XlChartType
        Dim region As String
        Dim c As Excel.Range
        Dim lTitle As String
        Dim lUnit As String
        Dim lComment As String
        Dim numCarScale As Double
        Dim numFormat As String
        Dim generateChart As Boolean
        Dim answerFile As String
        Dim emptyCell As Excel.Range
        Dim addRows As Integer
        Dim conStr As String
        Dim dHeight As Integer
        Dim dWidth As Integer
        Dim all_regions As String
        Const nRows As Integer = 24
        Const nCols As Integer = 8

        'set global variables: end-of-period, and step-size
        eop = CInt(TimeHorizonBox.Text)
        If stepCheckBox.Checked Then
            timestep = timestep * 2
        End If

        'add a sheet
        oSheet = excelWorkBook.Sheets.Add()
        oSheet.Select()

        'name the new sheet, subtitle in caseNameLong
        If StrComp(caseNameLong, "") = 0 Then
            sheetName = caseName
        Else
            sheetName = caseNameLong
        End If

        'solve naming conflicts by adding "(n)", n = 1,2,3,...
        If WorksheetExists(sheetName) Then
            i = 1
            Do
                sheetName2 = sheetName & " (" & CStr(i) & ")"
                i = i + 1
            Loop Until Not WorksheetExists(sheetName2)
            excelWorkBook.Sheets(sheetName).Name = sheetName2
        End If
        oSheet.Name = sheetName


        'set local variables
        generateChart = chartCheckBox.Checked
        numCarScale = CDbl(numCarScaleBox.Text)
        answerFile = FilenameBox.Text

        'set some info on the new sheet
        With excelWorkBook.ActiveSheet
            .Range("A1").Value = "CaseName: " & caseName
            .Range("A2").Value = "Answer file: " & answerFileName
            .Range("A3").Value = Now
            .Range("A1:A3").Font.Bold = True
            dHeight = nRows * .Rows(1).RowHeight ' height of all charts
            dWidth = nCols * .Columns("A").Width ' width of all charts
        End With


        If selectRegions.Checked Then
            all_regions = InputBox("Which region (ALL = all regions)", "Select a region", "ALL")
        Else
            all_regions = "ALL"
        End If

        'open the connection to the ACCESS database
        con = New ADODB.Connection 'early binding
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & answerFileName
        con.Open(conStr)
        rs = New ADODB.Recordset 'early binding

        'excute the global and local queries
        For i = 1 To 100
            c = excelWorkBook.Sheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)
            If Not c Is Nothing Then
                lTitle = CStr(c.Offset(0, 1).Value)
                lUnit = CStr(c.Offset(1, 1).Value)
                lComment = CStr(c.Offset(2, 1).Value)
                numFormat = IIf(String.IsNullOrEmpty(c.Offset(4, 1).Value), "0.0", CStr(c.Offset(4, 1).Value))
                If Not String.IsNullOrEmpty(c.Offset(5, 1).Value) And
         Not String.IsNullOrEmpty(c.Offset(6, 1).Value) Then
                    oRange = excelWorkBook.Sheets("Sheet1").Range(c.Offset(5, 1), c.Offset(5, 1).End(XlDirection.xlDown))
                Else
                    oRange = c.Offset(5, 1)
                End If
                Select Case c.Offset(3, 1).Value
                    Case "LineMarkers"
                        lChartType = Excel.XlChartType.xlLineMarkers
                    Case "AreaStacked"
                        lChartType = Excel.XlChartType.xlAreaStacked
                    Case Else
                        lChartType = Excel.XlChartType.xlColumnStacked 'default
                End Select

                str = ""
                'Global
                If c.Offset(0, -2).Value = "*" Then
                    Select Case c.Value
                        Case 1 'Total Primary Energy by Type
                            str = SQLquery1(caseName, allRegions)
                        Case 2 'Total Primary Energy by Region
                            str = SQLquery2(caseName)
                        Case 3 'Emission by Region
                            str = SQLquery3(caseName)
                        Case 4 'CCS
                            str = SQLquery4(caseName)
                        Case 5 ' Fuels for Long-Range Cars
                            str = sqlQuery5(caseName, allRegions)
                        Case 6 'Fuels for Short-Range Cars
                            str = sqlQuery6(caseName, allRegions)
                        Case 7 'Fuels for Cars
                            str = sqlQuery7(caseName, allRegions)
                        Case 8 'Fuels in Transport
                            str = SQLquery8(caseName, allRegions)
                        Case 9 ' T2 Technologies
                            str = SQLquery9(caseName, allRegions)
                        Case 10 'Detailed Technology Mix of Mileage for Long-Range
                            str = sqlQuery10(caseName, allRegions)
                        Case 11 'Technology Mix of Mileage for Short-Range Cars
                            str = SQLquery11(caseName, allRegions)
                        Case 12 'Technology Mix of Mileage for Cars
                            str = SQLquery12(caseName, allRegions)
                        Case 13 'Technology Mix for new Long-Range Cars
                            str = SQLquery13(caseName, allRegions, numCarScale)
                        Case 14 'Technology Mix for new Short-Range Cars
                            str = SQLquery14(caseName, allRegions, numCarScale)
                        Case 15 'Technology Mix for new Cars
                            str = SQLquery15(caseName, allRegions, numCarScale)
                        Case 16 'Fuel Refining and Synthesis
                            str = sqlQuery16(caseName, allRegions)
                        Case 17 'Biofuel Production
                            str = SQLquery17(caseName, allRegions)
                        Case 18 'Hydrogen Production
                            str = SQLquery18(caseName, allRegions)
                        Case 19 'Hydrogen Production (Detail)
                            str = SQLquery19(caseName, allRegions)
                        Case 20 'Use of Hydrogen
                            str = SQLquery20(caseName, allRegions)
                        Case 21 'Electricity Production by Primary Energy
                            str = SQLquery21(caseName, allRegions)
                        Case 22 'Renewable Electricity Production
                            str = SQLquery22(caseName, allRegions)
                        Case 23 'Electricity CO2 Emission
                            str = SQLquery23(caseName, allRegions)
                        Case 24 'BEV CO2 Emission
                            str = SQLQuery24(caseName, allRegions)
                        Case 25 'Depletion of Oil Resources
                            str = sqlQuery25(caseName, allRegions, scenarioName)
                        Case 26 'Oil Production
                            str = sqlQuery26(caseName)
                        Case 27 'Gas Production
                            str = sqlQuery27(caseName)
                        Case 28 'Coal Production
                            str = sqlQuery28(caseName)
                        Case 29 'Biomass Production
                            str = sqlQuery29(caseName, allRegions)
                        Case 30 'Marginal Cost of Fuel (average)
                            str = sqlQuery30(caseName, allRegions)
                        Case 31 'Fuel in T1 and T3
                            str = sqlQuery31(caseName, allRegions)
                        Case 32 'Technologies in T1 and T3
                            str = sqlQuery32(caseName, allRegions)
                        Case 33 'Emissions from T2 and T4 by Region
                            str = sqlQuery33(caseName)
                        Case 34 'Detailed Fuels in Transport
                            str = sqlQuery34(caseName, allRegions)
                        Case 35 'Detailed Car Technology Cardinality, Diesel
                            str = SQLquery35(caseName, allRegions)
                        Case 36 'GDP Market Scenario by Region
                            str = SQLquery36(caseName)
                        Case 37 'GDP State Scenario by Region
                            str = SQLquery37(caseName)
                        Case 38 'Population Market Scenario by Region
                            str = SQLquery38(caseName)
                        Case 39 'Population State Scenario by Region
                            str = SQLquery39(caseName)
                        Case 40 'Fuels in Industry Thermal
                            str = SQLquery40(caseName, allRegions)
                        Case 41 'Fuels in Industry Specific
                            str = SQLquery41(caseName, allRegions)
                        Case 42 'Fuels in Feedstocks
                            str = SQLquery42(caseName, allRegions)
                        Case 43 'Fuels in Residential Commercial Thermal
                            str = SQLquery43(caseName, allRegions)
                        Case 44 'Fuels in Residential Commercial Specific
                            str = SQLquery44(caseName, allRegions)
                        Case 45 'Electricity Production from Coal and Gas
                            str = SQLquery45(caseName, allRegions)
                        Case 46 'Oil Production by Category
                            str = sqlQuery46(caseName, allRegions)
                        Case 47 'Gas Production by Category
                            str = sqlQuery47(caseName, allRegions)
                        Case 48 'Net Import by Region
                            str = sqlQuery48(caseName)
                        Case 49 'Trade by Region
                            str = sqlQuery49(caseName)
                        Case 50 'Depletion of Gas Resources
                            str = sqlQuery50(caseName, allRegions, scenarioName)
                        Case 51 'TFC by fuel type
                            str = SQLquery51(caseName, allRegions)
                        Case 52 'Electricity Capacity by Primary Energy
                            str = SQLquery52(caseName, allRegions)
                        Case 53 'Coal Net Import by Region
                            str = sqlQuery53(caseName)
                        Case 54 'Gas Net Import by Region
                            str = sqlQuery54(caseName)
                        Case 55 'Oil Net Import by Region
                            str = sqlQuery55(caseName)
                        Case 56 'Bio Net Import by Region
                            str = sqlQuery56(caseName)
                        Case 57 'TFC by sector
                            str = SQLquery57(caseName, allRegions)
                        Case 58 'TFC by region
                            str = SQLquery58(caseName)
                        Case 59 'Industry/Feedstock TFC by region
                            str = SQLquery59(caseName)
                        Case 60 'Res./Comm. TFC by region
                            str = SQLquery60(caseName)
                        Case 61 'Electricity TFC by region
                            str = SQLquery61(caseName)
                        Case 62 'TFC Feedstocks by Region
                            str = SQLquery62(caseName)
                        Case 63 'Fuels in Residential/Commercial Total
                            str = SQLquery63(caseName, allRegions)
                        Case 64 'Fuels in Industry Total
                            str = SQLquery64(caseName, allRegions)
                        Case 65 'Fuels in Industry and feedstocsk Total
                            str = SQLquery65(caseName, allRegions)
                        Case 66 'Electricity TFC by region
                            str = SQLquery66(caseName)
                        Case 67 'Number of cars per region
                            str = SQLquery67(caseName)
                        Case 68 'Total net imports by Region
                            str = sqlQuery68(caseName)
            'Case 69 'xxx
            '    str = sqlQuery69(caseName)
                        Case 70 'LNG Net Import by Region
                            str = sqlQuery70(caseName)
                        Case 71 'Hydrogen Net Import by Region
                            str = sqlQuery71(caseName)
                        Case 72 'Net Diesel-Import/Export by Region
                            str = sqlQuery72(caseName)
                        Case 73 'Net Gasoline Import/Export by Region
                            str = sqlQuery73(caseName)
                        Case 74 'Biodiesel Import/Export by Region
                            str = sqlQuery74(caseName)
                        Case 75 'Bio-Ethanol Import/Export by Region
                            str = sqlQuery75(caseName)
                        Case 76 'Biogas Import/Export by Region
                            str = sqlQuery76(caseName)
                        Case 77 'FT-Hydrogen Import/Export by Region
                            str = sqlQuery77(caseName)
                        Case 78 'FT-Green Import/Export by Region
                            str = sqlQuery78(caseName)
                        Case 79 'Green Hydrogen Import/Export by Region
                            str = sqlQuery79(caseName)
                        Case 80 'Hydrogen Import/Export by Region
                            str = sqlQuery80(caseName)
                        Case 81 'MeOH Import/Export by Region
                            str = sqlQuery81(caseName)
                    End Select
                    If str <> "" Then
                        Call importData(answerFileName:=answerFile, lQuery:=str, lTitle:=lTitle, lUnit:=lUnit,
                        lChart:=lComment, lRange:=oRange, lType:=lChartType,
                        generateChart:=generateChart, caseNameLong:=caseNameLong,
                        lTitleShort:="GLOBAL" & "_" & lTitle, numFormat:=numFormat, dH:=dHeight, dW:=dWidth)
                    End If

                    ' add additional empty ("'") rows
                    If CType(c.Offset(0, 2).Value, String) <> "" Then
                        emptyCell = excelWorkBook.ActiveSheet.Cells(65536, 17).End(Excel.XlDirection.xlUp)
                        addRows = c.Offset(0, 2).Value - (excelWorkBook.ActiveSheet.Range(emptyCell.End(Excel.XlDirection.xlUp), emptyCell).Rows.Count - 3)
                        If addRows > 0 Then
                            excelWorkBook.ActiveSheet.Range(emptyCell.Offset(1), emptyCell.Offset(addRows)).Value = "'"
                        End If
                    End If

                End If 'global is selected
                str = ""
                'Regional
                If c.Offset(0, -1).Value = "*" Then
                    For Each r In regionArray
                        If UCase(r) = UCase(all_regions) Or UCase(all_regions) = "ALL" Then
                            region = "'" & r & "'"
                            Select Case c.Value
                                Case 1 'Total Primary Energy by Type
                                    str = SQLquery1(caseName, region)
                'Case 2 'Total Primary Energy by Region
                '    str = SQLquery2(caseName)
                'Case 3 'Emission by Region
                '    str = SQLquery3(caseName)
                'Case 4 'CCS
                '    str = SQLquery4(caseName)
                                Case 5 ' Fuels for Long-Range Cars
                                    str = sqlQuery5(caseName, region)
                                Case 6 'Fuels for Short-Range Cars
                                    str = sqlQuery6(caseName, region)
                                Case 7 'Fuels for Cars
                                    str = sqlQuery7(caseName, region)
                                Case 8 'Fuels in Transport
                                    str = SQLquery8(caseName, region)
                                Case 9 ' T2 Technologies
                                    str = SQLquery9(caseName, region)
                                Case 10 'Detailed Technology Mix of Mileage for Long-Range
                                    str = sqlQuery10(caseName, region)
                                Case 11 'Technology Mix of Mileage for Short-Range Cars
                                    str = SQLquery11(caseName, region)
                                Case 12 'Technology Mix of Mileage for Cars
                                    str = SQLquery12(caseName, region)
                                Case 13 'Technology Mix for new Long-Range Cars
                                    str = SQLquery13(caseName, region, numCarScale)
                                Case 14 'Technology Mix for new Short-Range Cars
                                    str = SQLquery14(caseName, region, numCarScale)
                                Case 15 'Technology Mix for new Cars
                                    str = SQLquery15(caseName, region, numCarScale)
                                Case 16 'Fuel Refining and Synthesis
                                    str = sqlQuery16(caseName, region)
                                Case 17 'Biofuel Production
                                    str = SQLquery17(caseName, region)
                                Case 18 'Hydrogen Production
                                    str = SQLquery18(caseName, region)
                                Case 19 'Hydrogen Production (Detail)
                                    str = SQLquery19(caseName, region)
                                Case 20 'Use of Hydrogen
                                    str = SQLquery20(caseName, region)
                                Case 21 'Electricity Production by Primary Energy
                                    str = SQLquery21(caseName, region)
                                Case 22 'Renewable Electricity Production
                                    str = SQLquery22(caseName, region)
                                Case 23 'Electricity CO2 Emission
                                    str = SQLquery23(caseName, region)
                                Case 24 'BEV CO2 Emission
                                    str = SQLQuery24(caseName, region)
                                Case 25 'Depletion of Oil Resources
                                    str = sqlQuery25(caseName, region, scenarioName)
                'Case 26 'Oil Production
                '    str = SQLquery26(caseName)
                'Case 27 'Gas Production
                '   str = SQLquery27(caseName)
                'Case 28 'Coal Production
                '    str = SQLquery28(caseName)
                                Case 29 'Biomass Production
                                    str = sqlQuery29(caseName, region)
                                Case 30 'Marginal Cost of Fuel (average)
                                    str = sqlQuery30(caseName, region)
                                Case 31 'Fuel in T1 and T3
                                    str = sqlQuery31(caseName, region)
                                Case 32 'Oil/Electricity in Transport Sectors
                                    str = sqlQuery32(caseName, region)
                'Case 33 'Emissions from T2 and T4 by Region
                '    str = SQLquery33(caseName)
                                Case 34 'Detailed Fuels in Transport
                                    str = sqlQuery34(caseName, region)
                                Case 35 'Detailed Car Technology Cardinality, Diesel
                                    str = SQLquery35(caseName, region)
                'Case 36 'GDP Market Scenario by Region
                '    str = SQLquery36(caseName)
                'Case 37 'GDP State Scenario by Region
                '    str = SQLquery37(caseName)
                'Case 38 'GDP Market Scenario by Region
                '    str = SQLquery38(caseName)
                'Case 39 'GDP State Scenario by Region
                '    str = SQLquery39(caseName)
                                Case 40 'Fuels in Industry Thermal
                                    str = SQLquery40(caseName, region)
                                Case 41 'Fuels in Industry Specific
                                    str = SQLquery41(caseName, region)
                                Case 42 'Fuels in Feedstocks
                                    str = SQLquery42(caseName, region)
                                Case 43 'Fuels in Residential Commercial Thermal
                                    str = SQLquery43(caseName, region)
                                Case 44 'Fuels in Residential Commercial Specific
                                    str = SQLquery44(caseName, region)
                                Case 45 'Electricity Production from Coal and Gas
                                    str = SQLquery45(caseName, region)
                                Case 46 'Oil Production by Category
                                    str = sqlQuery46(caseName, region)
                                Case 47 'Gas Production by Category
                                    str = sqlQuery47(caseName, region)
                'Case 48 'Net Import by Region
                '    str = sqlQuery48(caseName)
                'Case 49 'Trade by Region
                '    str = sqlQuery49(caseName)
                                Case 50 'Depletion of Gas Resources
                                    str = sqlQuery50(caseName, region, scenarioName)
                                Case 51 'TFC by fuel type
                                    str = SQLquery51(caseName, region)
                                Case 52 'Electricity Capacity by Primary Energy
                                    str = SQLquery52(caseName, region)
                'Case 53 'Coal Net Import by Region
                '   str = sqlQuery53(caseName)
                'Case 54 'Gas Net Import by Region
                '   str = sqlQuery54(caseName)
                'Case 55 'Oil Net Import by Region
                '   str = sqlQuery55(caseName)
                'Case 56 'Bio Net Import by Region
                '   str = sqlQuery56(caseName)
                                Case 57 'TFC by sector
                                    str = SQLquery57(caseName, region)
                'Case 58 'TFC by region
                'str = SQLquery58(caseName)
                 'Case 59 'Industry/Feedstock TFC by region
                'str = SQLquery59(caseName)
                'Case 60 'Res./Comm. TFC by region
                'str = SQLquery60(caseName)
                'Case 61 'Electricity TFC by region
                'str = SQLquery61(caseName)
                'Case 62 'TFC Feedstocks by Region
                'str = SQLquery62(caseName)
                                Case 63 'Fuels in Residential/Commercial total by Region
                                    str = SQLquery63(caseName, region)
                                Case 64 'Fuels in Industry total by Region
                                    str = SQLquery64(caseName, region)
                                Case 65 'Fuels in Industry including feedstocks by Region
                                    str = SQLquery65(caseName, region)
                                Case 69 'Fuels in Power Generation by Region
                                    str = SQLquery69(caseName, region)
                                    'Case 70 'LNG Net Import by Region
                                    'str = sqlQuery70(caseName)
                                    'Case 71 'Hydrogen Net Import by Region
                                    'str = sqlQuery71(caseName)
                                    'Case 72 'Net Diesel-Import/Export by Region
                                    'str = sqlQuery72(caseName)
                                    'Case 73 'Net Gasoline Import/Export by Region
                                    'str = sqlQuery73(caseName)
                                    'Case 74 'Biodiesel Import/Export by Region
                                    'str = sqlQuery74(caseName)
                                    'Case 75 'Bio-Ethanol Import/Export by Region
                                    'str = sqlQuery75(caseName)
                                    'Case 76 'Biogas Import/Export by Region
                                    'str = sqlQuery76(caseName)
                                    'Case 77 'FT-Hydrogen Import/Export by Region
                                    'str = sqlQuery77(caseName)
                                    'Case 78 'FT-Green Import/Export by Region
                                    'str = sqlQuery78(caseName)
                                    'Case 79 'Green Hydrogen Import/Export by Region
                                    'str = sqlQuery79(caseName)
                                    'Case 80 'Hydrogen Import/Export by Region
                                    'str = sqlQuery80(caseName)
                                    'Case 81 'MeOH Import/Export by Region
                                    'str = sqlQuery81(caseName)
                            End Select
                            If str <> "" Then
                                Call importData(answerFileName:=answerFile, lQuery:=str, lTitle:=lTitle & ", Region " & r,
                            lUnit:=lUnit, lChart:=lComment, lRange:=oRange,
                            lType:=lChartType, generateChart:=generateChart,
                            caseNameLong:=caseNameLong, lTitleShort:=r & "_" & lTitle, numFormat:=numFormat, dH:=dHeight, dW:=dWidth)
                            End If

                            ' add additional empty ("'") rows
                            If CType(c.Offset(0, 2).Value, String) <> "" Then
                                emptyCell = excelWorkBook.ActiveSheet.Cells(65536, 17).End(Excel.XlDirection.xlUp)
                                addRows = c.Offset(0, 2).Value - (excelWorkBook.ActiveSheet.Range(emptyCell.End(Excel.XlDirection.xlUp), emptyCell).Rows.Count - 3)
                                If addRows > 0 Then
                                    excelWorkBook.ActiveSheet.Range(emptyCell.Offset(1), emptyCell.Offset(addRows)).Value = "'"
                                End If
                            End If

                        End If
                    Next r
                End If 'regional is selected
            End If 'chart/table number i is to be extracted
        Next i

        'close the connection to the ACCESS database
        rs = Nothing
        con.Close()
        con = Nothing

        'arrange charts
        Call ArrangeMyCharts()
        'put second axis
        If axisCheckBox.Checked Then
            Call SecondaryAxis()
        End If
        'apply style
        If Excel2003Style.Checked Then
            Call Excel2003Styles()
        End If

    End Sub




    'https://gist.github.com/mikelheim/9087511
    'takes 2D array of any type and returns a transposed 2D array of same type. Vb.Net Function

    Function MyTranspose(Of Array)(ByVal inArray As Array(,)) As Array(,)
        Dim x = CInt(inArray.GetLength(1))
        Dim y = CInt(inArray.GetLength(0))
        Dim outArray(x - 1, y - 1) As Array

        For i = 0 To x - 1
            For j = 0 To y - 1
                outArray(i, j) = inArray(j, i)
            Next
        Next
        MyTranspose = outArray
    End Function

    'Public Function MyTranspose(ByVal vArr As Object) As Object
    ' Dim lRow As Long
    'Dim lCol As Long
    'Dim vNewArray() As Object
    '
    'ReDim vNewArray(LBound(vArr, 2) To UBound(vArr, 2), LBound(vArr, 1) To UBound(vArr, 1))
    'For lRow = LBound(vArr, 1) To UBound(vArr, 1)
    'For lCol = LBound(vArr, 2) To UBound(vArr, 2)
    '           vNewArray(lCol, lRow) = vArr(lRow, lCol)
    'Next
    'Next
    '   MyTranspose = vNewArray
    'End Function

    Sub importData(answerFileName As String, lQuery As String, lTitle As String,
               lUnit As String, lChart As String, lRange As Range, lType As Excel.XlChartType,
               generateChart As Boolean, caseNameLong As String, lTitleShort As String, Optional numFormat As String = "0.0", Optional dH As Integer = 150, Optional dW As Integer = 200)


        Dim nextFreeRow As Excel.Range
        Dim resizedRange As Excel.Range
        Dim myValues As Object
        'Dim myValues2  As Variant
        Dim nRows As Integer
        Dim nCols As Integer
        Dim i As Integer
        Dim sFieldname As String
        Dim dataRng As Excel.Range

        'Excel 2007: Provider=Microsoft.ACE.OLEDB.12.0;
        ' Needs:
        ' Excel-> VBA Editor -> Tools -> References "Microsoft Active Data Objects 2.1 Library" (perhaps also: "...multidimensional 2.8 Library")

        nextFreeRow = excelApp.ActiveSheet.Cells(65536, 17).End(XlDirection.xlUp)
        'MsgBox ActiveSheet.Name
        'MsgBox nextFreeRow.Row
        'MsgBox nextFreeRow.Column
        nextFreeRow = nextFreeRow.Offset(RowOffset:=3)
        nextFreeRow.Value = lTitle & " (Units: " & lUnit & ")"
        nextFreeRow = nextFreeRow.Offset(RowOffset:=1)
        'MsgBox nextFreeRow.Row
        'MsgBox nextFreeRow.Column

        If speedImport Then


            'Dim con         As ADODB.Connection 'early binding
            ''Dim con         As Object 'late binding
            'Dim conStr   As String
            'Dim rs          As ADODB.Recordset 'early binding
            ''Dim rs          As Object 'late binding


            ''Set con = CreateObject("ADODB.connection") 'ADO; late binding
            'Set con = New ADODB.Connection 'early binding
            'conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & answerFileName
            'con.Open conStr
            ''Set rs = CreateObject("ADODB.recordset") 'late binding
            'Set rs = New ADODB.Recordset 'early binding
            'If rs IsNot Nothing Then MsgBox("rs is defined")
            'MsgBox(lQuery)
            With rs
                '.DataSource = lQuery
                '.ActiveConnection = con
                '.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
                .Open(lQuery, con, CursorTypeEnum.adOpenForwardOnly)
                myValues = .GetRows
                nRows = UBound(myValues, 1) + 1
                nCols = UBound(myValues, 2) + 1
                'MsgBox "Rows returned = " & nRows
                'MsgBox "Cols returned = " & nCols
                'MsgBox CStr(myValues)
                'nextFreeRow.CopyFromRecordset .DataSource
                'WorksheetFunction.Transpose(arr)


                For i = 1 To .Fields.Count - 1
                    sFieldname = .Fields(i).Name
                    nextFreeRow.Offset(ColumnOffset:=i).Value = sFieldname
                    nextFreeRow.Offset(ColumnOffset:=i).Font.Bold = True
                    'MsgBox sFieldname
                Next

                resizedRange = nextFreeRow.Resize(nCols + 1, nRows)
                nextFreeRow = nextFreeRow.Offset(RowOffset:=1)
                dataRng = nextFreeRow.Resize(nCols, nRows)
                'MsgBox "Rows in Range = " & dataRng.Rows.Count
                'MsgBox "Cols in Range = " & dataRng.Columns.Count
                'myValues2 = MyTranspose(myValues)
                dataRng.Value = MyTranspose(myValues)

                'MsgBox .Fields(0).Value
            End With
            ''Set thee cursor location.
            'rs.CursorLocation = 3 'adUseClient on early  binding
            'rs.CursorType = 1 'adOpenKeyset on early  binding
            'rs.Open lQuery, con
            ''Redim the table that will contain the filtered data.
            'ReDim myValues(rs.RecordCount)
            'If Not (rs.EOF And rs.BOF) Then
            ' rs.MoveFirst
            ' Dim dbcol As Integer
            ' dbcol = 0
            ' nextFreeRow.Value = rs(dbcol).Value
            rs.Close()
            'Set rs = Nothing
            'con.Close
            'Set con = Nothing

        Else

            Dim oQryTable As Object
            oQryTable = excelApp.ActiveSheet.QueryTables.Add(
   "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & answerFileName & ";",
   nextFreeRow, lQuery)
            '"Select Region, Item1, Item2, Item3, CaseName, Parameter, T2000 from tblTRESULTS")
            'make room for new rows
            oQryTable.RefreshStyle = XlCellInsertionMode.xlInsertEntireRows
            oQryTable.Refresh(False)

            resizedRange = oQryTable.ResultRange

            'MsgBox "Delete Query Table"
            oQryTable.Delete

        End If 'speedup or not

        If Len(resizedRange.Rows(resizedRange.Rows.Count).Cells(1, 1).Value) = 0 Then
            'MsgBox "Warning: Name of last data row is empty; the row will not be charted. Chart = " & lTitle
            resizedRange = resizedRange.Resize(resizedRange.Rows.Count - 1)
            'MsgBox resizedRange.Rows.Count
            'MsgBox resizedRange.Columns.Count
        End If
        'MsgBox "Finished Table!"



        Call permutate(resizedRange, lRange)

        If lookupCheckBox.Checked Then
            Call lookupCol(resizedRange, lRange, lTitleShort, caseNameLong)
        End If

        If generateChart Then
            If StrComp(caseNameLong, "") <> 0 Then
                lTitle = lTitle & Chr(10) & caseNameLong
            End If
            Call addChart(resizedRange, lTitle, lUnit, lRange, lType, lChart, dH, dW)
        End If

        'MsgBox "Finished Chart!"
        resizedRange.Offset(1).Resize(resizedRange.Rows.Count - 1).NumberFormat = numFormat

    End Sub


    Sub addChart(lRange As Range, lTitle As String, lUnit As String, lcolorRange As Range, lChartType As Excel.XlChartType, lComment As String, dH As Integer, dW As Integer)

        Dim chtChart As Chart
        Dim lSheet As String

        lSheet = excelApp.ActiveSheet.Name
        chtChart = excelApp.ActiveSheet.Shapes.addChart(10, 10, dW, dH).Chart
        'Set chtChart = chtChart.Location(WHERE:=xlLocationAsObject, Name:=lSheet)
        With chtChart
            .DisplayBlanksAs = XlDisplayBlanksAs.xlZero
            .ChartType = lChartType
            .SetSourceData(Source:=lRange, PlotBy:=XlRowCol.xlRows)
            .HasTitle = True
            'please comment the next line out if you are using Excel 2003
            '.SetElement(2)   ' msoElementChartTitleAboveChart CHANGE
            .HasTitle = True
            .ChartTitle.Text = lTitle
            .ChartTitle.Font.Bold = True
            .PlotArea.Interior.ColorIndex = XlColorIndex.xlColorIndexNone
            .PlotArea.Border.LineStyle = XlLineStyle.xlLineStyleNone
            .ChartGroups(1).GapWidth = 20
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).HasTitle = True
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Characters.Text = lUnit
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Font.Size = 11
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.Font.Size = 10
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.NumberFormatLinked = False
            '.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).TickLabels.numberFormat = "General"
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "0"
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.Font.Size = 10
            .Legend.Font.Size = 10
            .ChartTitle.Font.Size = 12
            .ChartArea.AutoScaleFont = False
            .Axes(XlAxisType.xlValue).MajorGridlines.Border.ColorIndex = 48
            .Axes(XlAxisType.xlValue).MajorGridlines.Border.Weight = XlBorderWeight.xlHairline
            .Axes(XlAxisType.xlValue).MajorGridlines.Border.LineStyle = XlLineStyle.xlDot
            If lComment <> vbNullString Then
                .Axes(XlAxisType.xlCategory).HasTitle = True
                .Axes(XlAxisType.xlCategory).AxisTitle.Text = lComment
                .Axes(XlAxisType.xlCategory).AxisTitle.Font.Size = 7
                .Axes(XlAxisType.xlCategory).AxisTitle.HorizontalAlignment = XlHAlign.xlHAlignLeft  'CHANGE!!!
            End If
            'The Parent property is used to set properties of
            'the ChartObject, size etc.
            'ActiveSheet.Range("A65536").End(xlUp).Row
            '.Parent.Top = lRange.Top
            '.Parent.Left = lRange.Columns(lRange.Columns.Count).Left + 9
            '.Parent.Left = lRange.Columns(1).Left
            'MsgBox lRange.Columns(lRange.Columns.Count).Left
            'MsgBox lRange.Columns(lRange.Columns.Count).Left + 9
            'MsgBox lRange.Columns.Count
            'MsgBox lRange.Left

            If .Legend.LegendEntries.Count = 1 Then
                .Legend.Delete()
            ElseIf .Legend.LegendEntries.Count > 15 Then
                .Legend.Font.Size = 8
            End If

        End With

        Call reColor(lcolorRange, chtChart)
        'Worksheets(lSheet).Activate

    End Sub



    'This example sums the data in the first column of query table one. The sum of the first column is displayed below the data range.
    'Set c1 = Sheets("sheet1").QueryTables(1).ResultRange.Columns(1)
    'c1.Name = "Column1"
    'c1.End(xlDown).Offset(2, 0).Formula = "=sum(Column1)"


    Sub ArrangeMyCharts(Optional nColumns As Integer = 2,
                    Optional shift As Integer = 0,
                    Optional nRows As Double = 24,
                    Optional nCols As Double = 8,
                    Optional reverse As Boolean = False)

        Dim iChart As Integer
        Dim nCharts As Integer
        Dim dTop As Integer
        Dim dLeft As Integer
        Dim dHeight As Integer
        Dim dWidth As Integer
        'Dim nColumns As Long

        With excelApp

            dHeight = nRows * .ActiveSheet.Rows(1).RowHeight ' height of all charts
            dWidth = nCols * .ActiveSheet.Columns("A").Width ' width of all charts
            dTop = 0      ' top of first row of charts
            dLeft = shift * dWidth ' left of first column of charts
            '    nColumns = 2   ' number of columns of charts
            nCharts = .ActiveSheet.ChartObjects.Count

        End With

        If reverse Then
            For iChart = 1 To nCharts
                With excelApp.ActiveSheet.ChartObjects(iChart)
                    .Height = dHeight
                    .Width = dWidth
                    .Top = dTop + Int(((nCharts - iChart + 1) - 1) / nColumns) * dHeight
                    .Left = dLeft + (((nCharts - iChart + 1) - 1) Mod nColumns) * dWidth
                End With
            Next
        Else
            For iChart = 1 To nCharts
                With excelApp.ActiveSheet.ChartObjects(iChart)
                    .Height = dHeight
                    .Width = dWidth
                    .Top = dTop + Int((iChart - 1) / nColumns) * dHeight
                    .Left = dLeft + ((iChart - 1) Mod nColumns) * dWidth
                End With
            Next
        End If

    End Sub

    Sub AutoScale_Off()

        Dim ws As Worksheet, co As ChartObject, i As Integer
        Dim ch As Chart

        With excelApp

            For Each ws In .ActiveWorkbook.Worksheets
                ' Go through each worksheet in the workbook
                For Each co In ws.ChartObjects
                    'In each chart turn the Auto Scale font feature off
                    i = i + 1
                    co.Chart.ChartArea.AutoScaleFont = False
                Next co
            Next ws
            For Each ch In .ActiveWorkbook.Charts
                'Go through each chart in the workbook
                ch.ChartArea.AutoScaleFont = False
                i = i + 1
            Next
            'MsgBox i & " charts have been altered"
            'Application.DisplayAlerts = True

        End With

    End Sub

    Sub permutate(r As Range, orderRange As Range)

        Dim TempArray As Object
        'Article ID: 169104 - Last Review: October 22, 2000 - Revision: 1.0
        'XL97: Run-Time Error Using Macro to Add Custom List
        ' -> workaround with a temp array

        With excelApp

            '    If orderRange.Rows.Count > 2 Then
            '    TempArray = orderRange
            '    excelApp.AddCustomList(ListArray:=CType(TempArray.Value, String()))
            '   r.Sort(Key1:=r.Cells(2, 1), Order1:=XlSortOrder.xlDescending, Header:=XlYesNoGuess.xlYes,
            'OrderCustom:= .CustomListCount + 1)
            '.DeleteCustomList(.CustomListCount)
            'End If

        End With

    End Sub

    Sub lookupCol(r As Range, lRange As Range, lTitleShort As String, caseNameLong As String)

        'caseNameLong is usually "", better to use caseName in Import

        Dim tagName
        Dim resizedRng As Excel.Range
        Dim regional As Boolean

        If r.Rows.Count - 1 > 0 Then
            'exclude first (header) row
            resizedRng = r.Offset(1, 0).Resize(r.Rows.Count - 1)

            If (StrComp(CStr(r.Cells(1, 1).Value), "Region") = 0) Then
                regional = True
                lTitleShort = Strings.Right(lTitleShort, Len(lTitleShort) - Len("GLOBAL"))
            Else
                regional = False
            End If

            For Each tagName In resizedRng.Columns(1).Cells
                If regional Then
                    tagName.Offset(0, -2).Value = tagName.Value & lTitleShort
                Else
                    tagName.Offset(0, -2).Value = lTitleShort & "_" & tagName.Value
                End If
                tagName.Offset(0, -1).Value = "'"
            Next tagName

        End If

    End Sub


    Sub reColor(ref As Range, cht As Chart)

        'Dim ser As Series
        Dim ser As Object     'CHANGE: Use late binding, because mso is not defined in early binding
        Dim refRow As Range
        Dim val2
        'Dim colArray

        For Each ser In cht.SeriesCollection
            With ser
                For Each refRow In ref
                    If (StrComp(CStr(refRow.Cells(1, 1).Value), .Name) = 0) Then
                        val2 = refRow.Cells(1, 3).Value
                        If cht.ChartType = XlChartType.xlColumnStacked Then
                            If RGBCheckBox.Checked And Not IsNothing(val2) Then  ' CHANGE, previous IsEmpty
                                'colArray = getRGB(CStr(val2))
                                'ser.Interior.Color = RGB(colArray(1), colArray(2), colArray(3))
                                .Interior.ColorIndex = CInt(val2)
                                If refRow.Cells(1, 4).Value = "s" Then
                                    .Fill.Patterned(16)    'msoPatternDarkUpwardDiagonal CHANGE
                                End If
                            Else
                                .Interior.ColorIndex = CInt(refRow.Cells(1, 2).Value)
                            End If
                        End If
                        If cht.ChartType = XlChartType.xlLineMarkers Then
                            If RGBCheckBox.Checked And Not IsNothing(val2) Then ' CHANGE, previous IsEmpty
                                'colArray = getRGB(CStr(val2))
                                'ser.Border.Color = RGB(colArray(1), colArray(2), colArray(3))
                                'ser.MarkerBackgroundColor = RGB(colArray(1), colArray(2), colArray(3))
                                'ser.MarkerForegroundColor = RGB(colArray(1), colArray(2), colArray(3))
                                .Border.ColorIndex = CInt(val2)
                                .MarkerBackgroundColorIndex = CInt(val2)
                                .MarkerForegroundColorIndex = CInt(val2)
                            Else
                                .Border.ColorIndex = CInt(refRow.Cells(1, 2).Value)
                                .MarkerBackgroundColorIndex = CInt(refRow.Cells(1, 2).Value)
                                .MarkerForegroundColorIndex = CInt(refRow.Cells(1, 2).Value)
                            End If
                        End If
                    End If
                Next refRow
            End With
        Next ser

    End Sub



End Class
