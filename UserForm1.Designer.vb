<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UserForm1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        CreateExcelFile = New Button()
        Label1 = New Label()
        selectAll = New Button()
        selectAllRegion = New Button()
        DeselectAll = New Button()
        colorButton = New Button()
        CopyDBandAggregateButton = New Button()
        Excel2003Style = New CheckBox()
        secondaryButton = New Button()
        arrangeButton = New Button()
        OpenExcelFile = New Button()
        ConvertToBWButton = New Button()
        equalizeButton = New Button()
        AvgButton = New Button()
        ResetColorButton = New Button()
        WECcolorButton = New Button()
        importButton = New Button()
        importAllButton = New Button()
        arrScenButton = New Button()
        revArrScnButton = New Button()
        stepCheckBox = New CheckBox()
        chartCheckBox = New CheckBox()
        fixedCheckBox = New CheckBox()
        axisCheckBox = New CheckBox()
        barrelCheckBox = New CheckBox()
        RGBCheckBox = New CheckBox()
        selectRegions = New CheckBox()
        lookupCheckBox = New CheckBox()
        ScreenUpdating = New CheckBox()
        unloadButton = New Button()
        CaseNameComboBox = New ComboBox()
        TimeHorizonBox = New TextBox()
        numCarScaleBox = New TextBox()
        ScenarioTextBox = New TextBox()
        TimeHorizonDBbox = New TextBox()
        FilenameBox = New TextBox()
        FilenameButton = New Button()
        ResultviewerFileNameBox = New TextBox()
        SuspendLayout()
        ' 
        ' CreateExcelFile
        ' 
        CreateExcelFile.Location = New Point(47, 12)
        CreateExcelFile.Name = "CreateExcelFile"
        CreateExcelFile.Size = New Size(150, 23)
        CreateExcelFile.TabIndex = 0
        CreateExcelFile.Text = "Create Excel File"
        CreateExcelFile.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(171, 58)
        Label1.Name = "Label1"
        Label1.Size = New Size(41, 15)
        Label1.TabIndex = 1
        Label1.Text = "Label1"
        ' 
        ' selectAll
        ' 
        selectAll.Location = New Point(31, 95)
        selectAll.Name = "selectAll"
        selectAll.Size = New Size(166, 23)
        selectAll.TabIndex = 2
        selectAll.Text = "Select All Queires (Global)"
        selectAll.UseVisualStyleBackColor = True
        ' 
        ' selectAllRegion
        ' 
        selectAllRegion.Location = New Point(24, 124)
        selectAllRegion.Name = "selectAllRegion"
        selectAllRegion.Size = New Size(173, 23)
        selectAllRegion.TabIndex = 3
        selectAllRegion.Text = "Select All Queries (Region)"
        selectAllRegion.UseVisualStyleBackColor = True
        ' 
        ' DeselectAll
        ' 
        DeselectAll.Location = New Point(24, 159)
        DeselectAll.Name = "DeselectAll"
        DeselectAll.Size = New Size(173, 23)
        DeselectAll.TabIndex = 4
        DeselectAll.Text = "Deselect All Queries"
        DeselectAll.UseVisualStyleBackColor = True
        ' 
        ' colorButton
        ' 
        colorButton.Location = New Point(24, 208)
        colorButton.Name = "colorButton"
        colorButton.Size = New Size(173, 23)
        colorButton.TabIndex = 5
        colorButton.Text = "Update Colors (E or F colm)"
        colorButton.UseVisualStyleBackColor = True
        ' 
        ' CopyDBandAggregateButton
        ' 
        CopyDBandAggregateButton.Location = New Point(24, 255)
        CopyDBandAggregateButton.Name = "CopyDBandAggregateButton"
        CopyDBandAggregateButton.Size = New Size(188, 23)
        CopyDBandAggregateButton.TabIndex = 6
        CopyDBandAggregateButton.Text = "Copy DB and Aggregate"
        CopyDBandAggregateButton.UseVisualStyleBackColor = True
        ' 
        ' Excel2003Style
        ' 
        Excel2003Style.AutoSize = True
        Excel2003Style.Checked = True
        Excel2003Style.CheckState = CheckState.Checked
        Excel2003Style.Location = New Point(47, 349)
        Excel2003Style.Name = "Excel2003Style"
        Excel2003Style.Size = New Size(137, 19)
        Excel2003Style.TabIndex = 7
        Excel2003Style.Text = "Excel 2003 chart style"
        Excel2003Style.UseVisualStyleBackColor = True
        ' 
        ' secondaryButton
        ' 
        secondaryButton.Location = New Point(288, 96)
        secondaryButton.Name = "secondaryButton"
        secondaryButton.Size = New Size(129, 23)
        secondaryButton.TabIndex = 8
        secondaryButton.Text = "Second Axis"
        secondaryButton.UseVisualStyleBackColor = True
        ' 
        ' arrangeButton
        ' 
        arrangeButton.Location = New Point(288, 130)
        arrangeButton.Name = "arrangeButton"
        arrangeButton.Size = New Size(129, 23)
        arrangeButton.TabIndex = 9
        arrangeButton.Text = "Arrange Charts"
        arrangeButton.UseVisualStyleBackColor = True
        ' 
        ' OpenExcelFile
        ' 
        OpenExcelFile.Location = New Point(3, 29)
        OpenExcelFile.Name = "OpenExcelFile"
        OpenExcelFile.Size = New Size(194, 23)
        OpenExcelFile.TabIndex = 10
        OpenExcelFile.Text = "Open Resultviewer.xlsx"
        OpenExcelFile.UseVisualStyleBackColor = True
        ' 
        ' ConvertToBWButton
        ' 
        ConvertToBWButton.Location = New Point(288, 159)
        ConvertToBWButton.Name = "ConvertToBWButton"
        ConvertToBWButton.Size = New Size(124, 23)
        ConvertToBWButton.TabIndex = 11
        ConvertToBWButton.Text = "Convert to BW"
        ConvertToBWButton.UseVisualStyleBackColor = True
        ' 
        ' equalizeButton
        ' 
        equalizeButton.Location = New Point(291, 200)
        equalizeButton.Name = "equalizeButton"
        equalizeButton.RightToLeft = RightToLeft.No
        equalizeButton.Size = New Size(126, 23)
        equalizeButton.TabIndex = 12
        equalizeButton.Text = "Equl. Axis (del 2nd)"
        equalizeButton.UseVisualStyleBackColor = True
        ' 
        ' AvgButton
        ' 
        AvgButton.Location = New Point(290, 229)
        AvgButton.Name = "AvgButton"
        AvgButton.Size = New Size(127, 23)
        AvgButton.TabIndex = 13
        AvgButton.Text = "avg. base"
        AvgButton.UseVisualStyleBackColor = True
        ' 
        ' ResetColorButton
        ' 
        ResetColorButton.Location = New Point(278, 266)
        ResetColorButton.Name = "ResetColorButton"
        ResetColorButton.Size = New Size(134, 23)
        ResetColorButton.TabIndex = 14
        ResetColorButton.Text = "default palette"
        ResetColorButton.UseVisualStyleBackColor = True
        ' 
        ' WECcolorButton
        ' 
        WECcolorButton.Location = New Point(266, 305)
        WECcolorButton.Name = "WECcolorButton"
        WECcolorButton.Size = New Size(151, 23)
        WECcolorButton.TabIndex = 15
        WECcolorButton.Text = "custom (tabl. on Sheet1)"
        WECcolorButton.UseVisualStyleBackColor = True
        ' 
        ' importButton
        ' 
        importButton.Location = New Point(506, 183)
        importButton.Name = "importButton"
        importButton.Size = New Size(75, 23)
        importButton.TabIndex = 16
        importButton.Text = "View Case"
        importButton.UseVisualStyleBackColor = True
        ' 
        ' importAllButton
        ' 
        importAllButton.Location = New Point(506, 236)
        importAllButton.Name = "importAllButton"
        importAllButton.Size = New Size(124, 23)
        importAllButton.TabIndex = 17
        importAllButton.Text = "View Cases (Batch)"
        importAllButton.UseVisualStyleBackColor = True
        ' 
        ' arrScenButton
        ' 
        arrScenButton.Location = New Point(509, 274)
        arrScenButton.Name = "arrScenButton"
        arrScenButton.Size = New Size(169, 23)
        arrScenButton.TabIndex = 18
        arrScenButton.Text = "Move charts to JAZZ sheet"
        arrScenButton.UseVisualStyleBackColor = True
        ' 
        ' revArrScnButton
        ' 
        revArrScnButton.Location = New Point(511, 308)
        revArrScnButton.Name = "revArrScnButton"
        revArrScnButton.Size = New Size(140, 23)
        revArrScnButton.TabIndex = 19
        revArrScnButton.Text = "Move charts back"
        revArrScnButton.UseVisualStyleBackColor = True
        ' 
        ' stepCheckBox
        ' 
        stepCheckBox.AutoSize = True
        stepCheckBox.Location = New Point(514, 66)
        stepCheckBox.Name = "stepCheckBox"
        stepCheckBox.Size = New Size(156, 19)
        stepCheckBox.TabIndex = 20
        stepCheckBox.Text = "2n steps only (5Y model)"
        stepCheckBox.UseVisualStyleBackColor = True
        ' 
        ' chartCheckBox
        ' 
        chartCheckBox.AutoSize = True
        chartCheckBox.Checked = True
        chartCheckBox.CheckState = CheckState.Checked
        chartCheckBox.Location = New Point(471, 99)
        chartCheckBox.Name = "chartCheckBox"
        chartCheckBox.Size = New Size(58, 19)
        chartCheckBox.TabIndex = 21
        chartCheckBox.Text = "charts"
        chartCheckBox.UseVisualStyleBackColor = True
        ' 
        ' fixedCheckBox
        ' 
        fixedCheckBox.AutoSize = True
        fixedCheckBox.Checked = True
        fixedCheckBox.CheckState = CheckState.Checked
        fixedCheckBox.Location = New Point(471, 128)
        fixedCheckBox.Name = "fixedCheckBox"
        fixedCheckBox.Size = New Size(133, 19)
        fixedCheckBox.TabIndex = 22
        fixedCheckBox.Text = "fixed Colors (not all)"
        fixedCheckBox.UseVisualStyleBackColor = True
        ' 
        ' axisCheckBox
        ' 
        axisCheckBox.AutoSize = True
        axisCheckBox.Location = New Point(471, 156)
        axisCheckBox.Name = "axisCheckBox"
        axisCheckBox.Size = New Size(90, 19)
        axisCheckBox.TabIndex = 23
        axisCheckBox.Text = "Second Axis"
        axisCheckBox.UseVisualStyleBackColor = True
        ' 
        ' barrelCheckBox
        ' 
        barrelCheckBox.AutoSize = True
        barrelCheckBox.Location = New Point(594, 154)
        barrelCheckBox.Name = "barrelCheckBox"
        barrelCheckBox.Size = New Size(113, 19)
        barrelCheckBox.TabIndex = 24
        barrelCheckBox.Text = "bboe (not mtoe)"
        barrelCheckBox.UseVisualStyleBackColor = True
        ' 
        ' RGBCheckBox
        ' 
        RGBCheckBox.AutoSize = True
        RGBCheckBox.Location = New Point(587, 131)
        RGBCheckBox.Name = "RGBCheckBox"
        RGBCheckBox.Size = New Size(163, 19)
        RGBCheckBox.TabIndex = 25
        RGBCheckBox.Text = "alternative colors (F colm)"
        RGBCheckBox.UseVisualStyleBackColor = True
        ' 
        ' selectRegions
        ' 
        selectRegions.AutoSize = True
        selectRegions.Location = New Point(673, 66)
        selectRegions.Name = "selectRegions"
        selectRegions.Size = New Size(127, 19)
        selectRegions.TabIndex = 26
        selectRegions.Text = "ask to select region"
        selectRegions.UseVisualStyleBackColor = True
        ' 
        ' lookupCheckBox
        ' 
        lookupCheckBox.AutoSize = True
        lookupCheckBox.Checked = True
        lookupCheckBox.CheckState = CheckState.Checked
        lookupCheckBox.Location = New Point(567, 99)
        lookupCheckBox.Name = "lookupCheckBox"
        lookupCheckBox.Size = New Size(109, 19)
        lookupCheckBox.TabIndex = 27
        lookupCheckBox.Text = "lookup Column"
        lookupCheckBox.UseVisualStyleBackColor = True
        ' 
        ' ScreenUpdating
        ' 
        ScreenUpdating.AutoSize = True
        ScreenUpdating.Location = New Point(672, 100)
        ScreenUpdating.Name = "ScreenUpdating"
        ScreenUpdating.Size = New Size(100, 19)
        ScreenUpdating.TabIndex = 28
        ScreenUpdating.Text = "screen update"
        ScreenUpdating.UseVisualStyleBackColor = True
        ' 
        ' unloadButton
        ' 
        unloadButton.Location = New Point(654, 395)
        unloadButton.Name = "unloadButton"
        unloadButton.Size = New Size(75, 23)
        unloadButton.TabIndex = 29
        unloadButton.Text = "Exit"
        unloadButton.UseVisualStyleBackColor = True
        ' 
        ' CaseNameComboBox
        ' 
        CaseNameComboBox.FormattingEnabled = True
        CaseNameComboBox.Location = New Point(283, 36)
        CaseNameComboBox.Name = "CaseNameComboBox"
        CaseNameComboBox.Size = New Size(121, 23)
        CaseNameComboBox.TabIndex = 30
        ' 
        ' TimeHorizonBox
        ' 
        TimeHorizonBox.Location = New Point(287, 73)
        TimeHorizonBox.Name = "TimeHorizonBox"
        TimeHorizonBox.Size = New Size(100, 23)
        TimeHorizonBox.TabIndex = 31
        TimeHorizonBox.Text = "2060"
        ' 
        ' numCarScaleBox
        ' 
        numCarScaleBox.Location = New Point(626, 183)
        numCarScaleBox.Name = "numCarScaleBox"
        numCarScaleBox.Size = New Size(100, 23)
        numCarScaleBox.TabIndex = 32
        numCarScaleBox.Text = "1.4"
        ' 
        ' ScenarioTextBox
        ' 
        ScenarioTextBox.Location = New Point(657, 224)
        ScenarioTextBox.Name = "ScenarioTextBox"
        ScenarioTextBox.Size = New Size(100, 23)
        ScenarioTextBox.TabIndex = 33
        ScenarioTextBox.Text = "BASE"
        ' 
        ' TimeHorizonDBbox
        ' 
        TimeHorizonDBbox.Location = New Point(38, 58)
        TimeHorizonDBbox.Name = "TimeHorizonDBbox"
        TimeHorizonDBbox.ReadOnly = True
        TimeHorizonDBbox.Size = New Size(100, 23)
        TimeHorizonDBbox.TabIndex = 34
        ' 
        ' FilenameBox
        ' 
        FilenameBox.Location = New Point(272, 9)
        FilenameBox.Name = "FilenameBox"
        FilenameBox.Size = New Size(320, 23)
        FilenameBox.TabIndex = 35
        ' 
        ' FilenameButton
        ' 
        FilenameButton.Location = New Point(627, 20)
        FilenameButton.Name = "FilenameButton"
        FilenameButton.Size = New Size(75, 23)
        FilenameButton.TabIndex = 36
        FilenameButton.Text = "..."
        FilenameButton.UseVisualStyleBackColor = True
        ' 
        ' ResultviewerFileNameBox
        ' 
        ResultviewerFileNameBox.Location = New Point(249, 391)
        ResultviewerFileNameBox.Name = "ResultviewerFileNameBox"
        ResultviewerFileNameBox.Size = New Size(280, 23)
        ResultviewerFileNameBox.TabIndex = 37
        ResultviewerFileNameBox.Text = "C:\Users\densing\Desktop\Resultviewer.xlsx"
        ' 
        ' UserForm1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(ResultviewerFileNameBox)
        Controls.Add(FilenameButton)
        Controls.Add(FilenameBox)
        Controls.Add(TimeHorizonDBbox)
        Controls.Add(ScenarioTextBox)
        Controls.Add(numCarScaleBox)
        Controls.Add(TimeHorizonBox)
        Controls.Add(CaseNameComboBox)
        Controls.Add(unloadButton)
        Controls.Add(ScreenUpdating)
        Controls.Add(lookupCheckBox)
        Controls.Add(selectRegions)
        Controls.Add(RGBCheckBox)
        Controls.Add(barrelCheckBox)
        Controls.Add(axisCheckBox)
        Controls.Add(fixedCheckBox)
        Controls.Add(chartCheckBox)
        Controls.Add(stepCheckBox)
        Controls.Add(revArrScnButton)
        Controls.Add(arrScenButton)
        Controls.Add(importAllButton)
        Controls.Add(importButton)
        Controls.Add(WECcolorButton)
        Controls.Add(ResetColorButton)
        Controls.Add(AvgButton)
        Controls.Add(equalizeButton)
        Controls.Add(ConvertToBWButton)
        Controls.Add(OpenExcelFile)
        Controls.Add(arrangeButton)
        Controls.Add(secondaryButton)
        Controls.Add(Excel2003Style)
        Controls.Add(CopyDBandAggregateButton)
        Controls.Add(colorButton)
        Controls.Add(DeselectAll)
        Controls.Add(selectAllRegion)
        Controls.Add(selectAll)
        Controls.Add(Label1)
        Controls.Add(CreateExcelFile)
        Name = "UserForm1"
        Text = "UserForm1"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents CreateExcelFile As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents selectAll As Button
    Friend WithEvents selectAllRegion As Button
    Friend WithEvents DeselectAll As Button
    Friend WithEvents colorButton As Button
    Friend WithEvents CopyDBandAggregateButton As Button
    Friend WithEvents Excel2003Style As CheckBox
    Friend WithEvents secondaryButton As Button
    Friend WithEvents arrangeButton As Button
    Friend WithEvents OpenExcelFile As Button
    Friend WithEvents ConvertToBWButton As Button
    Friend WithEvents equalizeButton As Button
    Friend WithEvents AvgButton As Button
    Friend WithEvents ResetColorButton As Button
    Friend WithEvents WECcolorButton As Button
    Friend WithEvents importButton As Button
    Friend WithEvents importAllButton As Button
    Friend WithEvents arrScenButton As Button
    Friend WithEvents revArrScnButton As Button
    Friend WithEvents stepCheckBox As CheckBox
    Friend WithEvents chartCheckBox As CheckBox
    Friend WithEvents fixedCheckBox As CheckBox
    Friend WithEvents axisCheckBox As CheckBox
    Friend WithEvents barrelCheckBox As CheckBox
    Friend WithEvents RGBCheckBox As CheckBox
    Friend WithEvents selectRegions As CheckBox
    Friend WithEvents lookupCheckBox As CheckBox
    Friend WithEvents ScreenUpdating As CheckBox
    Friend WithEvents unloadButton As Button
    Friend WithEvents CaseNameComboBox As ComboBox
    Friend WithEvents TimeHorizonBox As TextBox
    Friend WithEvents numCarScaleBox As TextBox
    Friend WithEvents ScenarioTextBox As TextBox
    Friend WithEvents TimeHorizonDBbox As TextBox
    Friend WithEvents FilenameBox As TextBox
    Friend WithEvents FilenameButton As Button
    Friend WithEvents ResultviewerFileNameBox As TextBox

End Class
