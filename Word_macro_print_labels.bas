Sub CreateCustomLabels()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Dim table As table
    Dim cell As cell
    Dim labelText As String
    Dim labelDoc As Document
    Dim rowIndex As Integer
    Dim lastRowIndex As Integer
    Dim labelCounter As Integer
    Dim labelRow As Integer
    Dim labelCol As Integer
    Dim totalLabels As Integer
    Dim templateCode As String
    Dim numRows As Integer
    Dim numColumns As Integer
    Dim rowHeight As Single
    Dim columnWidth As Single

    ' Define each column in the table. In this example, skip Column 1 and Column >7
    Dim column2, column3, column4, column5, column6, column7 As String

    ' Define 3 types of attributes that will be colored on the labels
    Dim type1Color As String
    Dim type2Color As String
    Dim type3Color As String

    ' Select active document
    Set doc = ActiveDocument

    ' Select document properties
    With doc.PageSetup
        .PaperSize = wdPaperLetter
        .LeftMargin = InchesToPoints(0.3)
    End With

    ' Prompt user for Template
    templateCode = InputBox("Enter Avery label template number (5167 or 6560):", "Avery label template number")
    If templateCode = "" Then
        MsgBox "Avery label template number is required.", vbExclamation
        GoTo ExitSub
    End If

    ' Prompt user for Project
    projectName = InputBox("Enter Project", "Project")
    If projectName = "" Then
        MsgBox "Project is required.", vbExclamation
        GoTo ExitSub
    End If

    ' Prompt user for Lot
    projectLot = InputBox("Enter Lot", "lot")
    If projectLot = "" Then
        MsgBox "Lot is required.", vbExclamation
        GoTo ExitSub
    End If

    ' Set label template settings based on the entered code
    Select Case templateCode
        Case "5167"
            numRows = 20
            numColumns = 4
            rowHeight = 0.5
            totalLabels = 400
            doc.PageSetup.PaperSize = wdPaperLetter

        Case "6560"
            numRows = 10
            numColumns = 3
            rowHeight = 1
            totalLabels = 400
            doc.PageSetup.PaperSize = wdPaperLetter

        Case Else
            MsgBox "Unsupported label template.", vbExclamation
            GoTo ExitSub
    End Select

    ' Prompt user for colors
    type1Color = InputBox("Enter color for 'Type 1' (Red, Blue, Green):", "Type 1 Color")
    type2Color = InputBox("Enter color for 'Type 2' (Red, Blue, Green):", "Type 2 Color")
    type3Color = InputBox("Enter color for 'Type 3' (Red, Blue, Green):", "Type 3 Color")

    ' Validate colors
    If Not IsValidColor(type1Color) Or Not IsValidColor(type2Color) Or Not IsValidColor(type3Color) Then
        MsgBox "Invalid color entered. Please use common color names like Red, Blue, Green.", vbExclamation
        Exit Sub
    End If

    ' Check if there is at least one table in the document
    If doc.Tables.Count = 0 Then
        MsgBox "The document does not contain any tables.", vbExclamation
        Exit Sub
    End If

    ' Reference the first table in the document
    Set table = doc.Tables(1)

    ' Create a new document for the labels
    Set labelDoc = Documents.Add
    SetupLabelTemplate labelDoc, numRows, numColumns, rowHeight

    ' Initialize row index and label counter
    rowIndex = 2 ' Start from the second row to skip the header
    lastRowIndex = table.Rows.Count
    labelCounter = 1

    ' Loop through each row in the table
    For rowIndex = 2 To lastRowIndex
        labelText = ""

        ' Get the contents of the specified columns (skip the first column)
        On Error Resume Next
        column2 = CleanCellText(table.cell(rowIndex, 2).Range.text)
        column3 = CleanCellText(table.cell(rowIndex, 3).Range.text) & " mL" 'Add or change units, if needed
        column4 = CleanCellText(table.cell(rowIndex, 4).Range.text)
        column5 = CleanCellText(table.cell(rowIndex, 5).Range.text) & " C" 'Add or change units, if needed
        column6 = CleanCellText(table.cell(rowIndex, 6).Range.text)
        column7 = CleanCellText(table.cell(rowIndex, 7).Range.text)
        On Error GoTo 0

        ' Concatenate the contents into the required format
        labelText = projectName & projectLot & vbCrLf & _
                    column2 & " " & column3 & vbCrLf & _
                    column4 & " " & column5 & vbCrLf & _
                    column6 & " " & column7

        ' Determine the position of the label in the table
        labelRow = ((labelCounter - 1) \ numColumns) + 1
        labelCol = ((labelCounter - 1) Mod numColumns) + 1

        ' Populate the label in the label template
        With labelDoc.Tables(1).cell(labelRow, labelCol).Range
            .text = labelText
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .Font.Name = "Arial"
            .Font.Size = 7
        End With

        ' Increment label counter
        labelCounter = labelCounter + 1

        ' Check if we need to add a new page for more labels
        If labelCounter > totalLabels Then
            labelDoc.Tables(1).Range.Collapse Direction:=wdCollapseEnd
            SetupLabelTemplate labelDoc, numRows, numColumns, rowHeight
            labelCounter = 1
        End If
    Next rowIndex

    ' Apply colors to the text in the labels
    ApplyColorsToLabels labelDoc, type1Color, type2Color, type3Color

    ' Get the current folder path
    currentFolder = ActiveDocument.Path
    If currentFolder = "" Then currentFolder = Application.Options.DefaultFilePath(wdDocumentsPath)

    ' Prompt user to select save location and file name
    Set FileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With FileDialog
        .Title = "Save Label Document Output"
        .InitialFileName = currentFolder & "\" & "Labels_" & projectName & "_" & projectLot & ".docx"
        If .Show = -1 Then
            FileName = .SelectedItems(1)
            'Save the label document
            labelDoc.SaveAs2 FileName

        Else
            MsgBox "File saving has been cancelled", vbExclamation
            GoTo ExitSub
        End If
    End With

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "An error occured: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Sub SetupLabelTemplate(doc As Document, numRows As Integer, numColumns As Integer, rowHeight As Single)
    doc.PageSetup.PageWidth = InchesToPoints(8.5)
    doc.PageSetup.PageHeight = InchesToPoints(11)
    doc.PageSetup.TopMargin = InchesToPoints(0.5)
    doc.PageSetup.BottomMargin = InchesToPoints(0.5)
    doc.PageSetup.LeftMargin = InchesToPoints(0.3)
    doc.PageSetup.RightMargin = InchesToPoints(0.25)
    doc.PageSetup.Orientation = wdOrientPortrait

    'Calculate label dimansions based on specified template
    Dim labelHeight As Single
    Dim labelWIdth As Single
    labelHeight = InchesToPoints(0.5)
    labelWIdth = InchesToPoints(2.1)

    'Calculate pitch dimensions
    Dim verticalPitch As Single
    Dim horizontalPitch As Single
    verticalPitch = InchesToPoints(0.5)
    horizontalPitch = InchesToPoints(2.05)

    'Calculate table dimensions based on label plus pitch dimensions
    Dim tableWidth As Single
    Dim tableHeight As Single
    tableWidth = numColumns * (labelWIdth + horizontalPitch) - horizontalPitch
    tableHeight = numRows * (labelHeight + verticalPitch) - verticalPitch

    With doc.Tables.Add(Range:=doc.Range, numRows:=numRows, numColumns:=numColumns)
        .TopPadding = 0
        .BottomPadding = 0
        .LeftPadding = 0
        .RightPadding = 0
        .Spacing = 0
        .AllowPageBreaks = False
        .PreferredWidthType = wdPreferredWidthAuto
        .Rows.Height = InchesToPoints(rowHeight)
        .Rows.HeightRule = wdRowHeightExactly
        .PreferredWidth = tableWidth
        .Rows.SetHeight rowHeight:=labelHeight, HeightRule:=wdRowHeightExactly
        .Columns.SetWidth columnWidth:=labelWIdth, RulerStyle:=wdAdjustNone

    End With
End Sub

Function CleanCellText(text As String) As String
    CleanCellText = Left(text, Len(text) - 2)
End Function

Function IsValidColor(color As String) As Boolean
    Select Case LCase(color)
        Case "red", "blue", "green"
            IsValidColor = True
        Case Else
            IsValidColor = False
    End Select
End Function

Sub ApplyColorsToLabels(labelDoc As Document, type1Color As String, type2Color As String, type3Color As String)
    Dim labelRange As Range
    Dim para As Paragraph

    For Each para In labelDoc.Paragraphs
        Set labelRange = para.Range
        ApplyColorToWord labelRange, "Type 1", type1Color
        ApplyColorToWord labelRange, "Type 2", type2Color
        ApplyColorToWord labelRange, "Type 3", type3Color
    Next para
End Sub

Sub ApplyColorToWord(rng As Range, word As String, color As String)
    Dim startPos As Long
    Dim endPos As Long
    Dim wordRange As Range

    startPos = InStr(1, rng.text, word, vbTextCompare)
    Do While startPos > 0
        endPos = startPos + Len(word) - 1
        Set wordRange = rng.Duplicate
        wordRange.Start = rng.Start + startPos - 1
        wordRange.End = rng.Start + endPos
        Select Case LCase(color)
            Case "red"
                wordRange.Font.color = RGB(255, 0, 0)
            Case "blue"
                wordRange.Font.color = RGB(0, 0, 255)
            Case "green"
                wordRange.Font.color = RGB(0, 255, 0)
            ' Add more colors as needed
            Case Else
                wordRange.Font.color = wdColorAutomatic
        End Select
        startPos = InStr(endPos + 1, rng.text, word, vbTextCompare)
    Loop
End Sub