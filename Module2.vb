Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Imports Windows.Win32.System


Public Module Module2

    'Prepare active chart for export
    Sub prepareChart()

        'Dim ws As Worksheet
        'Dim co As ChartObject

        If ActiveChart Is Nothing Then
            MsgBox("Please click (activate)  first a chart.")
            Exit Sub
        End If

        With ActiveChart
            .ChartArea.Border.LineStyle = XlLineStyle.xlLineStyleNone
            .ChartArea.Interior.ColorIndex = xlColorIndexNone
            '.Axes(xlValue, xlPrimary).TickLabels.Font.Size = _
            '.Axes(xlValue, xlPrimary).TickLabels.Font.Size + 1
            '.Axes(xlCategory, xlPrimary).TickLabels.Font.Size = _
            '.Axes(xlCategory, xlPrimary).TickLabels.Font.Size + 1
            'If .HasAxis(xlValue, xlSecondary) Then
            '    .Axes(xlValue, xlSecondary).TickLabels.Font.Size = _
            '    .Axes(xlValue, xlSecondary).TickLabels.Font.Size + 1
            'End If
        End With

    End Sub


    'Increase: Ticklabels (x,y)
    Sub Increase_Ticklabels()

        Dim ws As Worksheet
        Dim co As ChartObject

        ws = ActiveWorkbook.ActiveSheet
        For Each co In ws.ChartObjects
            With co.Chart
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.Font.Size =
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.Font.Size + 1
                .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.Font.Size =
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.Font.Size + 1
                If .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlSecondary) Then
                    .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.Font.Size =
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.Font.Size + 1
                End If
            End With
        Next co

    End Sub

    'Decrease: Ticklabels (x,y)
    Sub Decrease_Ticklabels()

        Dim ws As Worksheet
        Dim co As ChartObject

        ws = ActiveWorkbook.ActiveSheet
        For Each co In ws.ChartObjects
            With co.Chart
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.Font.Size =
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.Font.Size - 1
                .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.Font.Size =
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.Font.Size - 1
                If .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlSecondary) Then
                    .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.Font.Size =
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.Font.Size - 1
                End If
            End With
        Next co

    End Sub

    'Increase: Titles, Axestitles, Legend
    Sub Increase_Titles_Legend()

        Dim ws As Worksheet
        Dim co As ChartObject

        ws = ActiveWorkbook.ActiveSheet
        For Each co In ws.ChartObjects
            With co.Chart
                .ChartTitle.Font.Size = co.Chart.ChartTitle.Font.Size + 1
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Font.Size =
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Font.Size + 1
                If .HasLegend Then
                    .Legend.Font.Size = co.Chart.Legend.Font.Size + 1
                End If
                If .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlSecondary) Then
                    .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Font.Size =
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Font.Size + 1
                End If
            End With
        Next co

    End Sub

    'Decrease: Titles, Axestitles, Legend
    Sub Decrease_Titles_Legend()

        Dim ws As Worksheet
        Dim co As ChartObject

        ws = ActiveWorkbook.ActiveSheet
        For Each co In ws.ChartObjects
            With co.Chart
                .ChartTitle.Font.Size = co.Chart.ChartTitle.Font.Size - 1
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Font.Size =
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Font.Size - 1
                If .HasLegend Then
                    .Legend.Font.Size = co.Chart.Legend.Font.Size - 1
                End If
                If .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlSecondary) Then
                    .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Font.Size =
                .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Font.Size - 1
                End If
            End With
        Next co

    End Sub





    Sub SecondaryAxis()

        Dim ws As Worksheet
        Dim co As ChartObject
        Dim n As Integer
        Dim f As Double
        Dim s As String
        Dim s2 As String
        Dim ax As Axis

        'For Each ws In ActiveWorkbook.Worksheets
        ws = ActiveWorkbook.ActiveSheet
        ' Go through each worksheet in the workbook
        For Each co In ws.ChartObjects
            With co.Chart
                'MsgBox .ChartTitle.Text
                If .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlPrimary) = True Then
                    s = ""
                    If .Axes(XlAxisType.xlValue).HasTitle Then
                        s = .Axes(XlAxisType.xlValue).AxisTitle.Text
                    End If
                    If ((InStr(s, "EJ") > 0) Or (InStr(s, "TW") > 0)) Then
                        For Each ax In .Axes
                            If ax.AxisGroup = XlAxisGroup.xlSecondary Then
                                ax.HasTitle = False
                                ax.Delete()
                            End If
                        Next ax
                        'MsgBox .SeriesCollection.Count
                        .SeriesCollection.NewSeries
                        n = .SeriesCollection.Count
                        'MsgBox n
                        .SeriesCollection(n).Values = "={0.000001}"
                        .SeriesCollection(n).AxisGroup = XlAxisGroup.xlSecondary
                        '.HasAxis(XlAxisType.xlCategory, XlAxisGroup.xlPrimary) = True
                        '.HasAxis(XlAxisType.xlCategory, XlAxisGroup.xlSecondary) = False
                        '.HasAxis(XlAxisType.xlValue, XlAxisGroup.xlPrimary) = True
                        .HasAxis(XlAxisType.xlValue, XlAxisGroup.xlSecondary) = True
                        .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).MinimumScale = 0
                        .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).MaximumScale = 1
                        '.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
                        '.Axes(xlCategory, xlSecondary).CategoryType = xlAutomatic
                        .SeriesCollection(n).AxisGroup = XlAxisGroup.xlSecondary
                        'MsgBox .SeriesCollection(n).AxisGroup
                        f = 23.88
                        If StrComp(s, "EJ") = 0 Then
                            s2 = "mtoe"
                            If UserForm1.barrelCheckBox.Checked Then
                                f = 0.18
                                s2 = "bboe"
                            End If
                        Else
                            s2 = "mtoe/y"
                            If UserForm1.barrelCheckBox.Checked Then
                                f = 0.18
                                s2 = "bboe"
                            End If
                        End If
                        If StrComp(s, "TWh/y") = 0 Then
                            '1000 TWh = 3.6 EJ
                            f = 0.0036
                            s2 = "EJ/y"
                        End If
                        If InStr(.ChartTitle.Text, "Oil Production") > 0 Then
                            f = 1 / 0.0061 / 365 '(assumption)
                            s2 = "mio bbl/d"
                        End If
                        If InStr(.ChartTitle.Text, "Gas Production") > 0 Then
                            f = 1 / 0.03472 '(IEA)
                            s2 = "bcm/y"
                        End If
                        If InStr(.ChartTitle.Text, "Coal Production") > 0 Then
                            '1 EJ = 34.1 mio t SKE (BMWI.de)
                            f = 34.1
                            s2 = "mio t SKE/y"
                        End If

                        .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).MaximumScale = f * (.Axes(XlAxisType.xlValue).MaximumScale)
                        .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).HasTitle = True
                        .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Text = s2
                        '.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.NumberFormatLinked = False
                        '.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).TickLabels.numberFormat = "General"
                        If .HasLegend Then
                            n = .Legend.LegendEntries.Count
                            If n > 0 Then
                                If n > 18 Then
                                    MsgBox("More than 18 legend entries, you may delete dummy entry manually in " & .ChartTitle.Text)
                                Else
                                    .Legend.LegendEntries(n).Delete
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next co
        'MsgBox "Finished adding second axes."
    End Sub




    Sub DeleteAllCharts()

        ActiveSheet.ChartObjects.Delete

    End Sub

    ' This procedure creates a new temporary toolbar.
    ' Create a new floating toolbar and make it visible.
    Sub AddNewToolBar()

        Dim ComBar As CommandBar, ComBarContrl As CommandBarControl
        On Error GoTo ErrorHandler
        On Error Resume Next

        'Delete the toolbar if it already exists
        CommandBars("My Toolbar").Delete
        ComBar = CommandBars.Add(Name:="My Toolbar", Position:=
    msoBarTop, Temporary:=True)
        ComBar.Visible = True

        ' Button: Panel
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "Panel"
            .Style = msoButtonCaption
            .TooltipText = "Control Panel for Import from Answer"
            .OnAction = "Show_Panel"
        End With

        ' Button: Print
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            '.FaceId = 1000 chooses an icon
            .Caption = "Pdf setup"
            .Style = msoButtonCaption
            .TooltipText = "Charts print setup for current page"
            .OnAction = "PrintSetup"
        End With

        ' Button: Scroll home
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "Up to A1"
            .Style = msoButtonCaption
            .TooltipText = "Scrolls all sheets to A1"
            .OnAction = "homeScroll"
        End With

        ' Button: Scroll down
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "Down Pg"
            .Style = msoButtonCaption
            .TooltipText = "Scrolls all sheets down"
            .OnAction = "scrollDown"
        End With

        ' Button: Increase title and legend
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "+Titles"
            .Style = msoButtonCaption
            .TooltipText = "Larger Titles and Legend"
            .OnAction = "Increase_Titles_Legend"
        End With

        ' Button: Decrease title and legend
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "-Titles"
            .Style = msoButtonCaption
            .TooltipText = "Smaller Titles and Legend"
            .OnAction = "Decrease_Titles_Legend"
        End With

        ' Button: Increase ticklabels
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "+Ticks"
            .Style = msoButtonCaption
            .TooltipText = "Larger Ticklabels"
            .OnAction = "Increase_Ticklabels"
        End With

        ' Button: Decrease ticklabels
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "-Ticks"
            .Style = msoButtonCaption
            .TooltipText = "Smaller Ticklabels"
            .OnAction = "Decrease_Ticklabels"
        End With

        ' Button: Remove background on active chart
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "Strip"
            .Style = msoButtonCaption
            .TooltipText = "Prepare selected chart for ppt"
            .OnAction = "prepareChart"
        End With

        ' Button: Delete all charts on active sheets
        ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
        With ComBarContrl
            .Caption = "Del charts"
            .Style = msoButtonCaption
            .TooltipText = "Delete all charts on active sheets"
            .OnAction = "DeleteAllCharts"
        End With


        Exit Sub
ErrorHandler:
        MsgBox("Error " & Err.Number & vbCr & Err.Description)
        Exit Sub
    End Sub


    'Delete the toolbar called by workbook close
    Sub DeleteToolbar()

        On Error Resume Next
        CommandBars("My Toolbar").Delete

    End Sub

    'pdf print setup on active sheet
    Sub PrintSetup(Optional x As Integer = 24, Optional y As Integer = 3)

        'x: chart height (unit: 1 row)
        'y: number of rows of charts on a page to print

        Dim i As Integer
        Dim j As Integer

        Application.ScreenUpdating = False

        With ActiveSheet

            i = .ChartObjects(.ChartObjects.Count).BottomRightCell.Row
            .PageSetup.PrintArea = "$A$1:$P$" & i
            .ResetAllPageBreaks
            ActiveWindow.View = xlPageBreakPreview
            On Error Resume Next
            .VPageBreaks(1).DragOff xlToRight, 1
    For j = y * x To i - 1 Step y * x
                'On Error Resume Next
                .HPageBreaks.Add.Range("A" & j + 1)
            Next j

        End With

        Application.ScreenUpdating = True



    End Sub


    Sub homeScroll()

        Dim savedSht As Worksheet
        Dim sht As Worksheet

        savedSht = ActiveSheet

        For Each sht In ActiveWorkbook.Worksheets
            sht.Activate()
            Application.Goto(Reference:=Range("A1"), Scroll:=True)
        Next sht

        savedSht.Activate()

    End Sub

    Sub scrollDown(Optional x As Integer = 48)

        'x: number of rows to scroll down

        Dim savedSht As Worksheet
        Dim sht As Worksheet

        savedSht = ActiveSheet

        For Each sht In ActiveWorkbook.Worksheets
            sht.Activate()
            Application.Goto(Reference:=ActiveCell.Offset(x), Scroll:=True)
        Next sht

        savedSht.Activate()

    End Sub


    Sub EqualizeAxes()
        'equalize y-axis and delete second axis on sheets JAZZ, ROCK, SYMPH

        Dim ws As Worksheet
        Dim i As Integer
        Dim chartCnt As Integer
        Dim maxVal(1 To 1000) As Double
        'Dim maxCharts As Integer
        Dim sh As Variant
        Dim sheetName As Variant

        sh = Array("JAZZ", "ROCK", "SYMPH")

        'by cycling through the sheets, get max value for each chart
        For i = LBound(maxVal) To UBound(maxVal)
            maxVal(i) = 0
            For Each sheetName In sh
                ws = ActiveWorkbook.Sheets(sheetName)
                chartCnt = ws.ChartObjects.Count
                If chartCnt >= 1 And i <= chartCnt Then
                    With ws.ChartObjects(i).Chart
                        'delete second y-axis
                        If .Axes.Count = 3 Then
                            If .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).HasTitle Then
                                .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).AxisTitle.Delete
                            End If
                            .Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary).Delete
                        End If
                        ' update previous maximum value
                        maxVal(i) = WorksheetFunction.Max(maxVal(i), .Axes(XlAxisType.xlValue).MaximumScale)
                        .Axes(XlAxisType.xlValue).MinimumScale = 0
                    End With
                End If
            Next sheetName
        Next i

        'set max value for each chart
        For Each sheetName In sh
            ws = ActiveWorkbook.Sheets(sheetName)
            With ws
                chartCnt = .ChartObjects.Count
                For i = 1 To chartCnt
                    .ChartObjects(i).Chart.Axes(XlAxisType.xlValue).MaximumScale = maxVal(i)
                Next i
            End With
        Next sheetName

    End Sub


    Sub arrChartsScen()
        'move charts on sheets ROCK, SYMPH to JAZZ

        Dim n1 As Long
        Dim n2 As Long
        Dim n3 As Long
        Dim sh As Variant
        Dim s As Variant
        Dim i As Integer
        Dim ch As ChartObject

        sh = Array("JAZZ", "ROCK", "SYMPH")

        For Each s In sh
            If Not WorksheetExists(s) Then
                MsgBox("Sheet " & s & "does not exist. Cannot copy. Exiting.")
                Exit Sub
            End If
        Next s

        n1 = Worksheets(sh(0)).ChartObjects.Count
        n2 = Worksheets(sh(1)).ChartObjects.Count
        n3 = Worksheets(sh(2)).ChartObjects.Count
        If (n2 <> n3) Or (n1 <> n3) Or (n1 <> n2) Then
            MsgBox("Different number of charts on " &
 Worksheets(sh(0)).Name & " / " & Worksheets(sh(1)).Name & " / " & Worksheets(sh(2)).Name &
 " (" & n1 & " / " & n2 & " / " & n3 & "). Taking minimum.")
            n1 = WorksheetFunction.Min(n1, n2, n3)
        End If

        ' If there are no charts, do not move
        If n1 = 0 Then
            MsgBox("At least one scenario sheet has no charts. No charts to put side-by-side. Exiting.")
            Exit Sub
        End If

        If Not UserForm1.selectRegions.Checked Then
            Call UserForm1.OnStart()
        End If

        ' On the JAZZ sheet, arrange charts in single column to get more horizontal space
        Worksheets(sh(0)).Select
        ArrangeMyCharts(nColumns:=1)


        For i = 1 To 2
            Worksheets(sh(i)).Select
            ArrangeMyCharts(nColumns:=1, shift:=i)     ' shift 1 or 2 chart-width to right
            'Move charts ot JAZZ
            For Each ch In Worksheets(sh(i)).ChartObjects
                ch.Chart.Location(xlLocationAsObject, sh(0))
            Next ch
        Next i

        If Not UserForm1.selectRegions.Checked Then
            Call UserForm1.OnEnd()
        End If

    End Sub


    Sub ReverseArrChartsScen()

        ' Move charts back from JAZZ to ROCK and SYMPH sheets

        Dim sh As Variant
        Dim n As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim idx As Integer

        sh = Array("JAZZ", "ROCK", "SYMPH")

        ' If there are remaining charts in ROCK or SYMPH, do not move charts back
        If Worksheets(sh(1)).ChartObjects.Count > 0 Or Worksheets(sh(2)).ChartObjects.Count > 0 Then
            MsgBox("There are existing charts in " & sh(1) & " or " & sh(2) & ". Hence, charts are not moved back. Exiting.")
            Exit Sub
        End If


        With Worksheets(sh(0))

            k = .ChartObjects.Count
            n = Application.WorksheetFunction.Floor(k / 3, 1)
            If (k Mod 3) > 0 Then
                MsgBox("The number " & k & " of charts on " & sh(0) & " is NOT a multiple of 3.")
            End If
            MsgBox("Moving " & n & " charts back to SYMPH and ROCK from JAZZ sheet. Number of charts currently on JAZZ:" & k)

            If Not UserForm1.selectRegions.Checked Then
                Call UserForm1.OnStart()
            End If

            ' Move charts back to ROCK and SYMPH (in correct order)
            For i = 1 To 2 Step 1
                idx = k - i * n + 1
                For j = 1 To n
                    .ChartObjects(idx).Chart.Location Where:=xlLocationAsObject, Name:=sh(3 - i)
            ' Chart object number is decremeted for higher chart object numbers, hence idx = const.
                Next j
            Next i
            .Select
            ArrangeMyCharts nColumns:=2

End With

        For i = 1 To 2
            Worksheets(sh(i)).Select
            ArrangeMyCharts nColumns:=2 ', reverse:=True 'account for reversed chart index
        Next i

        If Not UserForm1.selectRegions.Checked Then
            Call UserForm1.OnEnd()
        End If

    End Sub


    Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean

        On Error Resume Next
        WorksheetExists = (ActiveWorkbook.Sheets(WorksheetName).Name <> "")
        On Error GoTo 0

    End Function

    Sub CopyDBandAggregate(Optional answerFileName As String = "M:\Martin\GMM\GMM 2013\GMM_Test.mdb")

        Dim dbConnectStr As String
        Dim cnt As ADODB.Connection
        Dim i As Integer
        Dim copyFileName As String

        'newRegions(3) As String
        'newRegions = Array("ASIA", "WEUR", "AFRICA")
        'SSAFRICA
        'MENA
        'LAC
        'NAM
        'EUROPE
        'CASIA
        'EASIA
        'PASIA

        Dim sqlQuery() As Variant
        sqlQuery = Array(
"ALTER TABLE tblTRESULTS DROP CONSTRAINT tblITEMStblTRESULTS",
"ALTER TABLE tblTRESULTS DROP CONSTRAINT tblITEMStblTRESULTS1",
"ALTER TABLE tblTRESULTS DROP CONSTRAINT tblITEMStblTRESULTS2",
"ALTER TABLE tblTRESULTS DROP CONSTRAINT tblITEMStblTRESULTS3",
"ALTER TABLE tblTRESULTS DROP CONSTRAINT tblITEMStblTRESULTS4",
"ALTER TABLE tblTRESULTS DROP CONSTRAINT tblITEMStblTRESULTS5",
"DROP INDEX PrimaryKey ON tblTRESULTS",
"INSERT INTO tblRegions VALUES ('@','MENA', 'Middle East and North Africa',1, ' ')",
"INSERT INTO tblRegions VALUES ('@','NAM', 'North America',1, ' ')",
"INSERT INTO tblRegions VALUES ('@','EUROPE', 'Europe',1, ' ')",
"INSERT INTO tblRegions VALUES ('@','CENTASIA', 'Central South Asia',1, ' ')",
"INSERT INTO tblRegions VALUES ('@','EASTASIA', 'East Asia',1, ' ')",
"INSERT INTO tblRegions VALUES ('@','PACIASIA', 'Southeast Asia Pacific',1, ' ')",
"UPDATE tblTRESULTS SET Region = 'LAC' WHERE Region IN ('BRAZIL')",
"UPDATE tblTRESULTS SET Region = 'MENA' WHERE Region IN ('GCC', 'MIDEAST', 'NAFRICA')",
"UPDATE tblTRESULTS SET Region = 'NAM' WHERE Region IN ('CANMEX', 'USA')",
"UPDATE tblTRESULTS SET Region = 'EUROPE' WHERE Region IN ('EU31', 'EEUR', 'RUSSIA')",
"UPDATE tblTRESULTS SET Region = 'CENTASIA' WHERE Region IN ('CENASIA', 'INDIA')",
"UPDATE tblTRESULTS SET Region = 'EASTASIA' WHERE Region IN ('CHINAREG', 'JPKRTW')",
"UPDATE tblTRESULTS SET Region = 'PACIASIA' WHERE Region IN ('ASIAPAC', 'AUSNZL')"
)

        copyFileName = Left(answerFileName, Len(answerFileName) - 4) & "_aggregated.mdb"
        MsgBox "new DB file name:" & copyFileName & ". Now copying ANSWER DB..."
FileCopy answerFileName, copyFileName
MsgBox "Copying finished. Removing MS Access DB Constraints and updating region names..."

dbConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & copyFileName & ";"

        cnt = New ADODB.Connection
        With cnt
            .Open(dbConnectStr)

            For i = LBound(sqlQuery, 1) To UBound(sqlQuery, 1)
                'MsgBox sqlQuery(i)
                .Execute sqlQuery(i)
    Next i
        End With
Set cnt = Nothing
MsgBox("Finished. You may select now the copied and updated DB.")

    End Sub

    Sub Dummy()
        Call CopyDBandAggregate()
    End Sub


    Sub BW()

        Dim i As Integer
        Dim ser As Series

        i = 1
        For Each ser In ActiveChart.SeriesCollection
            With ser
                Select Case i
                    Case 1, 9
                        'black with white dots
                        .Fill.ForeColor.SchemeColor = 1
                        .Fill.BackColor.SchemeColor = 2
                        '.Fill.Visible = True
                        .Fill.Patterned(msoPattern80Percent)
                    Case 2, 10
                        'diagional down
                        .Fill.ForeColor.SchemeColor = 1
                        .Fill.BackColor.SchemeColor = 2
                        '.Fill.Visible = True
                        .Fill.Patterned(msoPatternLightDownwardDiagonal)
                    Case 3, 11
                        'grey
                        .Interior.ColorIndex = 15
            '.Interior.Pattern = xlSolid
                    Case 4, 12
                        'white with black dots
                        .Fill.ForeColor.SchemeColor = 1
                        .Fill.BackColor.SchemeColor = 2
                        '.Fill.Visible = True
                        .Fill.Patterned(msoPattern20Percent)
                    Case 5, 13
                        'black
                        .Interior.ColorIndex = 1
             '.Interior.Pattern = xlSolid
                    Case 6, 14
                        'dense black dots
                        .Fill.ForeColor.SchemeColor = 1
                        .Fill.BackColor.SchemeColor = 2
                        '.Fill.Visible = True
                        .Fill.Patterned(msoPattern50Percent)
                    Case 7, 15
                        'white
                        .Interior.ColorIndex = 2
            '.Interior.Pattern = xlSolid
                    Case 8, 16
                        'diagonal up
                        .Fill.ForeColor.SchemeColor = 1
                        .Fill.BackColor.SchemeColor = 2
                        '.Fill.Visible = True
                        .Fill.Patterned(msoPatternLightUpwardDiagonal)
                End Select
                i = i + 1
            End With
        Next ser

    End Sub


    Sub Excel2003Styles()

        Dim i As Integer
        Dim ser As Series
        'Dim ch As Excel.Chart
        Dim co As Excel.ChartObject

        For Each co In excelActiveSheet.ChartObjects

            If co.Chart.HasLegend Then
                With co.Chart.Legend
                    .Format.Line.Visible = msoTrue
                    .Format.Line.ForeColor.ObjectThemeColor = 13 'msoThemeColorText1
                    For i = 1 To .LegendEntries.Count
                        If co.Chart.ChartType <> XlChartType.xlLineMarkers Then
                            .LegendEntries(i).LegendKey.Select
                            Selection.Format.Line.Visible = msoTrue
                            Selection.Format.Line.ForeColor.ObjectThemeColor = 13 'msoThemeColorText1
                        Else
                            .LegendEntries(i).LegendKey.Select
                            Selection.Format.Line.Visible = msoTrue
                            Selection.Format.Line.Weight = 1
                        End If
                    Next i
                    .Format.TextFrame2.TextRange.Font.Name = "Arial"
                End With
            End If

            For Each ser In ActiveChart.SeriesCollection
                With ser
                    If co.Chart.ChartType = XlChartType.xlColumnStacked Then
                        .Format.Line.Visible = msoTrue
                        .Format.Line.ForeColor.ObjectThemeColor = 13 'msoThemeColorText1
                    ElseIf co.Chart.ChartType = XlChartType.xlLineMarkers Then
                        'MsgBox .Format.Line.Weight
                        .Format.Line.Weight = 1
                    ElseIf co.Chart.ChartType = XlChartType.xlAreaStacked Then
                        .Format.Line.Visible = msoTrue
                        .Format.Line.Weight = 0.75
                    End If
                End With
            Next ser

            With co.Chart.Axes(xlValue)
                .Format.Line.ForeColor.ObjectThemeColor = 13 'msoThemeColorText1
                .AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Arial"
                .TickLabels.Font.Name = "Arial"
            End With

            With co.Chart.Axes(xlCategory)
                .Format.Line.ForeColor.ObjectThemeColor = 13 'msoThemeColorText1
                .TickLabels.Font.Name = "Arial"
                If .HasTitle Then
                    .AxisTitle.Format.TextFrame2.TextRange.Font.Name = "Arial"
                End If
            End With

            co.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = "Arial"


        Next co
    End Sub


    Function getRGB(c As String) As Integer()

        Dim colArray As Variant
        Dim lenColArray As Integer
        Dim colRet(3) As Integer

        colArray = Split(c, ",")
        lenColArray = UBound(colArray) - LBound(colArray) + 1
        If lenColArray <> 3 Then
            'Excel color Index
            MsgBox("The RGB color string = " & c & " has " & lenColArray & " and not 3 as length")
        Else
            For i = 1 To 3
                colRet(i) = colArray(LBound(colArray) + i - 1)
            Next i
        End If

        getRGB = colRet

    End Function


    Sub ResetColorPalette()
        ' Sub IS NO LONGER USED
        excelWorkbook.Colors(1) = RGB(0, 0, 0)
        excelWorkBook.Colors(2) = RGB(255, 255, 255)
        ThisWorkbook.Colors(3) = RGB(255, 0, 0)
        ThisWorkbook.Colors(4) = RGB(0, 255, 0)
        ThisWorkbook.Colors(5) = RGB(0, 0, 255)
        ThisWorkbook.Colors(6) = RGB(255, 255, 0)
        ThisWorkbook.Colors(7) = RGB(255, 0, 255)
        ThisWorkbook.Colors(8) = RGB(0, 255, 255)
        ThisWorkbook.Colors(9) = RGB(128, 0, 0)
        ThisWorkbook.Colors(10) = RGB(0, 128, 0)
        ThisWorkbook.Colors(11) = RGB(0, 0, 128)
        ThisWorkbook.Colors(12) = RGB(128, 128, 0)
        ThisWorkbook.Colors(13) = RGB(128, 0, 128)
        ThisWorkbook.Colors(14) = RGB(0, 128, 128)
        ThisWorkbook.Colors(15) = RGB(192, 192, 192)
        ThisWorkbook.Colors(16) = RGB(128, 128, 128)
        ThisWorkbook.Colors(17) = RGB(153, 153, 255)
        ThisWorkbook.Colors(18) = RGB(153, 51, 102)
        ThisWorkbook.Colors(19) = RGB(255, 255, 204)
        ThisWorkbook.Colors(20) = RGB(204, 255, 255)
        ThisWorkbook.Colors(21) = RGB(102, 0, 102)
        ThisWorkbook.Colors(22) = RGB(255, 128, 128)
        ThisWorkbook.Colors(23) = RGB(0, 102, 204)
        ThisWorkbook.Colors(24) = RGB(204, 204, 255)
        ThisWorkbook.Colors(25) = RGB(0, 0, 128)
        ThisWorkbook.Colors(26) = RGB(255, 0, 255)
        ThisWorkbook.Colors(27) = RGB(255, 255, 0)
        ThisWorkbook.Colors(28) = RGB(0, 255, 255)
        ThisWorkbook.Colors(29) = RGB(128, 0, 128)
        ThisWorkbook.Colors(30) = RGB(128, 0, 0)
        ThisWorkbook.Colors(31) = RGB(0, 128, 128)
        ThisWorkbook.Colors(32) = RGB(0, 0, 255)
        ThisWorkbook.Colors(33) = RGB(0, 204, 255)
        ThisWorkbook.Colors(34) = RGB(204, 255, 255)
        ThisWorkbook.Colors(35) = RGB(204, 255, 204)
        ThisWorkbook.Colors(36) = RGB(255, 255, 153)
        ThisWorkbook.Colors(37) = RGB(153, 204, 255)
        ThisWorkbook.Colors(38) = RGB(255, 153, 204)
        ThisWorkbook.Colors(39) = RGB(204, 153, 255)
        ThisWorkbook.Colors(40) = RGB(255, 204, 153)
        ThisWorkbook.Colors(41) = RGB(51, 102, 255)
        ThisWorkbook.Colors(42) = RGB(51, 204, 204)
        ThisWorkbook.Colors(43) = RGB(153, 204, 0)
        ThisWorkbook.Colors(44) = RGB(255, 204, 0)
        ThisWorkbook.Colors(45) = RGB(255, 153, 0)
        ThisWorkbook.Colors(46) = RGB(255, 102, 0)
        excelWorkBook.Colors(47) = RGB(102, 102, 153)
        Workbook.Colors(48) = RGB(150, 150, 150)
        Workbook.Colors(49) = RGB(0, 51, 102)
        Workbook.Colors(50) = RGB(51, 153, 102)
        Workbook.Colors(51) = RGB(0, 51, 0)
        Workbook.Colors(52) = RGB(51, 51, 0)
        Workbook.Colors(53) = RGB(153, 51, 0)
        Workbook.Colors(54) = RGB(153, 51, 102)
        Workbook.Colors(55) = RGB(51, 51, 153)
        Workbook.Colors(56) = RGB(51, 51, 51)

    End Sub

    Sub WECcolorPalette()
        ' SUB is no longer used
        'Regions
        Workbook.Colors(3) = RGB(10, 63, 102)
        Workbook.Colors(4) = RGB(33, 77, 129)
        Workbook.Colors(5) = RGB(33, 107, 167)
        Workbook.Colors(6) = RGB(48, 130, 191)
        Workbook.Colors(7) = RGB(76, 156, 217)
        Workbook.Colors(8) = RGB(117, 184, 235)
        Workbook.Colors(9) = RGB(166, 206, 255)
        Workbook.Colors(10) = RGB(3, 38, 64)
        Workbook.Colors(11) = RGB(211, 219, 230)
        'Primary
        Workbook.Colors(12) = RGB(213, 0, 50)
        'Fossil
        Workbook.Colors(13) = RGB(166, 88, 43)
        Workbook.Colors(14) = RGB(214, 154, 57)
        Workbook.Colors(15) = RGB(235, 202, 132)
        'Renewables
        Workbook.Colors(16) = RGB(188, 242, 55)
        Workbook.Colors(17) = RGB(102, 145, 22)
        Workbook.Colors(18) = RGB(51, 88, 0)
        Workbook.Colors(19) = RGB(160, 153, 88)
        Workbook.Colors(20) = RGB(181, 167, 69)
        Workbook.Colors(21) = RGB(90, 195, 121)
        Workbook.Colors(22) = RGB(46, 170, 151)
        Workbook.Colors(23) = RGB(104, 36, 75)
        Workbook.Colors(24) = RGB(157, 36, 75)
        ' Jazz
        Workbook.Colors(25) = RGB(253, 237, 199)
        ' Symphony
        Workbook.Colors(26) = RGB(228, 241, 245)
        ' Fossil
        Workbook.Colors(27) = RGB(235, 132, 3)
        ' Renewables
        Workbook.Colors(28) = RGB(137, 203, 57)
        ' Other
        Workbook.Colors(29) = RGB(104, 36, 75)
        ' remaining: standard Excel colors
        Workbook.Colors(30) = RGB(255, 0, 0)
        Workbook.Colors(31) = RGB(0, 255, 0)
        Workbook.Colors(32) = RGB(0, 0, 255)
        Workbook.Colors(33) = RGB(255, 0, 0)
        Workbook.Colors(34) = RGB(255, 0, 0)
        Workbook.Colors(35) = RGB(255, 0, 0)
        Workbook.Colors(36) = RGB(255, 0, 0)
        Workbook.Colors(37) = RGB(255, 0, 0)
        Workbook.Colors(38) = RGB(255, 0, 0)
        Workbook.Colors(39) = RGB(255, 0, 0)
        Workbook.Colors(40) = RGB(255, 0, 0)
        Workbook.Colors(41) = RGB(255, 0, 0)
        Workbook.Colors(42) = RGB(255, 0, 0)
        Workbook.Colors(43) = RGB(255, 0, 0)
        Workbook.Colors(44) = RGB(255, 0, 0)
        Workbook.Colors(45) = RGB(255, 0, 0)
        Workbook.Colors(46) = RGB(255, 0, 0)
        Workbook.Colors(47) = RGB(255, 0, 0)
        Workbook.Colors(48) = RGB(255, 0, 0)
        Workbook.Colors(49) = RGB(255, 0, 0)
        Workbook.Colors(50) = RGB(255, 0, 0)
        Workbook.Colors(51) = RGB(255, 0, 0)
        Workbook.Colors(52) = RGB(255, 0, 0)
        Workbook.Colors(53) = RGB(255, 0, 0)
        Workbook.Colors(54) = RGB(255, 0, 0)
        Workbook.Colors(55) = RGB(255, 0, 0)
        Workbook.Colors(56) = RGB(255, 0, 0)

    End Sub



    Sub setColorPalette()

        Dim rgbList As Range
        Dim c As Range
        Dim palette
        'Dim index As Integer
        Dim RGBcol() As Integer
        Dim i As Integer
  
 Set c = Worksheets("Sheet1").Range("A1:Z100").Find("Palette", LookIn:=xlValues, LookAt:=xlWhole)
 If c Is Nothing Then
            MsgBox("Could not find the 'RGB->colorindex'-list in Range A1:Z100. The list should have 'Palette' as header.")
        Else
    Set rgbList = Range(c.Offset(1, 0), c.End(xlDown))
    palette = ActiveWorkbook.Colors
            For Each r In rgbList.Rows
                i = CInt(r.Cells(1, 1).Value)
                If i < 1 Or i > 57 Then
                    MsgBox "ExcelColor Index " & i & " must be in range 1-57"
     Else
                    RGBcol = getRGB(CStr(r.Cells(1, 2).Value))
                    palette(i) = RGB(RGBcol(1), RGBcol(2), RGBcol(3))
                    r.Cells(1, 2).Interior.ColorIndex = i
                End If
            Next r
            ActiveWorkbook.Colors = palette
        End If
        MsgBox "Excel's color palette updated from Table 'Palette' on 'Sheet1'."

End Sub




End Module

