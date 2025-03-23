Imports System.Xml
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Imports Windows.Win32.System


Public Module Module3

    Sub avgBaseYear()
        ' average the base year numbers in JAZZ, SYMPH, ROCK

        Dim tags As Range, c As Range
        Dim rg1 As Range    ' lookup range on 1st sheet
        Dim rg2 As Range    ' lookup range on 2nd sheet
        Dim sh(3) As String ' Array of Sheet names
        Dim s As String     ' value to lookup
        Dim ret1 As Object  ' return of Vlookup on 1st sheet
        Dim ret2 As Object  ' return of Vlookup on 2nd sheet
        Dim i As Double ' divide by i for average
        Dim n As Integer 'index of sheet

        sh = {"JAZZ", "ROCK", "SYMPH"}

        With UserForm1.excelWorkBook

            For n = 0 To 2

                tags = .Sheets(sh(n Mod 3)).Range("O1:O10000")
                rg1 = .Sheets(sh((n + 1) Mod 3)).Range("O1:R10000")
                rg2 = .Sheets(sh((n + 2) Mod 3)).Range("O1:R10000")

                For Each c In tags
                    If Not IsNothing(c.Value) Then  'VBA: IsEmpty
                        s = c.Value
                        i = 1
                        ret1 = UserForm1.excelApp.VLookup(s, rg1, 4, False)
                        If IsError(ret1) Then
                            ret1 = 0
                        Else : i = i + 1
                        End If
                        ret2 = UserForm1.excelApp.VLookup(s, rg2, 4, False)
                        If IsError(ret2) Then
                            ret2 = 0
                        Else : i = i + 1
                        End If
                        c.Offset(0, 1).Value = (ret1 + ret2 + c.Offset(0, 3).Value) / i
                    End If
                Next c
            Next n

            For n = 0 To 2
                rg1 = .Sheets(sh(n)).Range("P1:P10000")
                For Each c In rg1
                    If Not IsNothing(c.Value) Then 'VBA: IsEmpty
                        c.Offset(0, 2).Value = c.Value
                        c.Value = "'"
                    End If
                Next c
            Next n


        End With
    End Sub



    Sub SplitMENASheet1()
        '
        ' SplitMENA Macro (for scenarios sheets)
        '
        ' Start: Select a cell above (or above-left) of the rows to be splitted, e.g. select the cell "AUSNZL"
        '
        ' 1. The search for "EU31" extends to 10 rows below and 30 columns right
        ' 2. Insert "GCC" row
        ' 3. The search range for "MENA" extends to 6 rows below and 30 columns right
        ' 4. Insert "MIDEAST" and "NAFRICA" row
        '
        ' Keyboard Shortcut: Ctrl+m
        '
        'Dim resultCell As Range

        With UserForm1.excelApp

            .ActiveCell.Resize(10, 30).Select()
            .Selection.Find(What:="EU31", After:= .ActiveCell, LookIn:=XlFindLookIn.xlValues,
            LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext,
            MatchCase:=True, SearchFormat:=False).Select
            .ActiveCell.Offset(1).EntireRow.Insert(XlDirection.xlDown)
            .ActiveCell.Offset.EntireRow.Copy()
            .ActiveCell.Offset(1).EntireRow.PasteSpecial(XlPasteType.xlPasteValues)
            .CutCopyMode = False
            .ActiveCell.Offset(0, 3).Value = "GCC"
            'ActiveCell.Offset(0, 6).Value = ""
            'ActiveCell.Offset(0, 7).Value = ""
            'ActiveCell.Offset(0, 8).Value = ""

            .Resize(6, 30).Select
            .Selection.Find(What:="MENA", After:= .ActiveCell, LookIn:=XlFindLookIn.xlValues,
            LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext,
            MatchCase:=True, SearchFormat:=False).Select
            .ActiveCell.Offset(1).EntireRow.Insert(XlDirection.xlDown)
            .ActiveCell.Offset(1).EntireRow.Insert(XlDirection.xlDown)
            .ActiveCell.Offset.EntireRow.Copy()
            .ActiveCell.Offset(1).EntireRow.PasteSpecial(XlPasteType.xlPasteValues)
            .ActiveCell.Offset(1).EntireRow.PasteSpecial(XlPasteType.xlPasteValues)
            .ActiveCell.Offset(-1, 3).Value = "MIDEAST"
            .ActiveCell.Offset(0, 3).Value = "NAFRICA"
            'ActiveCell.Offset(-1, 6).Value = ""
            'ActiveCell.Offset(0, 6).Value = ""
            'ActiveCell.Offset(-1, 7).Value = ""
            'ActiveCell.Offset(0, 7).Value = ""
            'ActiveCell.Offset(-1, 8).Value = ""
            'ActiveCell.Offset(0, 8).Value = ""
            .CutCopyMode = False

        End With
    End Sub


End Module



