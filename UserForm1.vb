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
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel





Public Class UserForm1

    Dim regionArray() As String
    Dim allRegions As String

    Public con As ADODB.Connection 'early binding
    Public rs As ADODB.Recordset 'early binding

    Public excelApp As Excel.Application
    Public excelWorkBook As Excel.Workbook
    Public excelWorkSheets As Excel.Worksheets
    Public excelWorkSheetOne As Excel.Worksheet
    Public excelActiveSheet As Excel.Worksheet


    Sub OpenExcelFile_Click(sender As Object, e As EventArgs) Handles OpenExcelFile.Click

        'Dim fn As String

        'Dim excelApp = New Excel.Application
        excelWorkBook = excelApp.Workbooks.Open(ResultviewerFileNameBox.Text)
        excelWorkSheets = excelWorkBook.Worksheets()
        excelWorkSheetOne = excelWorkBook.Worksheets("Sheet1")
        excelActiveSheet = excelApp.ActiveSheet
        'display the cells value B2
        MsgBox(excelWorkSheetOne.Cells(2, 2).value)

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

        Call BW()

    End Sub

    Private Sub selectAllRegion_Click() Handles selectAllRegion.Click

        Dim i As Integer
        Dim c As Excel.Range

        'Iterate to search for chart numbers in column C, and then put "*" (if not "NA") in column B
        For i = 1 To 100
            c = excelWorkSheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
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
            c = excelWorkSheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
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
            c = excelWorkSheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
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
            val = excelWorkSheets("Sheet1").Cells(i, 5).Value
            val2 = excelWorkSheets("Sheet1").Cells(i, 6).Value
            valChck = excelWorkSheets("Sheet1").Cells(i, 3).Value
            'If i < 20 Then
            '   MsgBox val2
            'End If
            If Not wf.IsNumber(valChck) Then
                If RGBCheckBox.Checked Then
                    If wf.IsNumber(val2) Then
                        If (0 < val2 And val2 < 58) Then
                            excelWorkSheets("Sheet1").Cells(i, 6).Interior.ColorIndex = CInt(val2)
                        Else
                            MsgBox("ColorIndex " & val2 & " of custom palette is out of range.")
                        End If
                    ElseIf Not String.IsNullOrEmpty(val2) Then
                        MsgBox("ColorIndex " & val2 & " is not a single, integer number.")
                    End If
                ElseIf wf.IsNumber(val) Then
                    If (0 < val And val < 58) Then
                        excelWorkSheets("Sheet1").Cells(i, 5).Interior.ColorIndex = CInt(val)
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
        Call ResetColorPalette()
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

        c = excelWorkSheets("Sheet1").Range("A1:Z100").Find("Batch", LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)
        If c Is Nothing Then
            MsgBox("Could not find list of cases in Range A1:Z100. The list as 'Batch' as header.")
        Else
            caseList = excelWorkSheets("Sheet1").Range(c.Offset(1, 0), c.End(XlDirection.xlDown))
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
        oSheet = excelWorkSheets.Add()
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
            excelWorkSheets(sheetName).Name = sheetName2
        End If
        oSheet.Name = sheetName


        'set local variables
        generateChart = chartCheckBox.Checked
        numCarScale = CDbl(numCarScaleBox.Text)
        answerFile = FilenameBox.Text

        'set some info on the new sheet
        With excelActiveSheet
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
            c = excelWorkSheets("Sheet1").Range("C1:C20000").Find(i, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)
            If Not c Is Nothing Then
                lTitle = CStr(c.Offset(0, 1).Value)
                lUnit = CStr(c.Offset(1, 1).Value)
                lComment = CStr(c.Offset(2, 1).Value)
                numFormat = IIf(String.IsNullOrEmpty(c.Offset(4, 1).Value), "0.0", CStr(c.Offset(4, 1).Value))
                If Not String.IsNullOrEmpty(c.Offset(5, 1).Value) And
         Not String.IsNullOrEmpty(c.Offset(6, 1).Value) Then
                    oRange = excelWorkSheets("Sheet1").Range(c.Offset(5, 1), c.Offset(5, 1).End(XlDirection.xlDown))
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
                        emptyCell = excelActiveSheet.Cells(65536, 17).End(Excel.XlDirection.xlUp)
                        addRows = c.Offset(0, 2).Value - (excelActiveSheet.Range(emptyCell.End(Excel.XlDirection.xlUp), emptyCell).Rows.Count - 3)
                        If addRows > 0 Then
                            excelActiveSheet.Range(emptyCell.Offset(1), emptyCell.Offset(addRows)).Value = "'"
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
                                emptyCell = excelActiveSheet.Cells(65536, 17).End(Excel.XlDirection.xlUp)
                                addRows = c.Offset(0, 2).Value - (excelActiveSheet.Range(emptyCell.End(Excel.XlDirection.xlUp), emptyCell).Rows.Count - 3)
                                If addRows > 0 Then
                                    excelActiveSheet.Range(emptyCell.Offset(1), emptyCell.Offset(addRows)).Value = "'"
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









End Class
