Sub CreateCustomLabels()
    On Error GoTo ErrorHandler
    Dim doc As Document
    Dim table As table
    Dim cell As cell
    Dim labelText As String
    Dim labelDoc As Document
    Dim rowIndex As Integer
    Dim lastRowIndex As Integer
    Dim volume, temperature, group, comment1, comment2, comment3 As String
    Dim labelCounter As Integer
    Dim labelRow As Integer
    Dim labelCol As Integer
    Dim totalLabels As Integer
    Dim templateCode As String
    Dim numRows As Integer
    Dim numColumns As Integer
    Dim rowHeight As Single
    Dim columnWidth As Single
    Dim activeColor As String
    Dim reagentColor As String
    Dim corrosiveColor As String

    ' Prompt user for Avery label template code
    templateCode = InputBox("Enter Avery label template code (e.g., 6560, 5160, etc.):", "Label Template Code")
    If templateCode = "" Then
        MsgBox "Template code is required.", vbExclamation
        Exit Sub
    End If

    ' Set label template settings based on the entered code
    Select Case templateCode
        Case "6560"
            numRows = 10
            numColumns = 3
            rowHeight = 1
            columnWidth = 2.625
            totalLabels = 30
        Case "5160"
            numRows = 10
            numColumns = 3
            rowHeight = 1
            columnWidth = 2.625
            totalLabels = 30
        ' Add more cases for other Avery templates as needed
        Case Else
            MsgBox "Unsupported label template code.", vbExclamation
            Exit Sub
    End Select

    ' Prompt user for colors
    activeColor = InputBox("Enter color for 'Active' (e.g., Red, Blue, Green):", "Active Color")
    reagentColor = InputBox("Enter color for 'Reagent' (e.g., Red, Blue, Green):", "Reagent Color")
    corrosiveColor = InputBox("Enter color for 'Corrosive' (e.g., Red, Blue, Green):", "Corrosive Color")

    ' Validate colors
    If Not IsValidColor(activeColor) Or Not IsValidColor(reagentColor) Or Not IsValidColor(corrosiveColor) Then
        MsgBox "Invalid color entered. Please use common color names like Red, Blue, Green.", vbExclamation
        Exit Sub
    End If

    ' Reference the active document
    Set doc = ActiveDocument

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
        volume = CleanCellText(table.cell(rowIndex, 2).Range.text)
        temperature = CleanCellText(table.cell(rowIndex, 3).Range.text)
        group = CleanCellText(table.cell(rowIndex, 4).Range.text)
        comment1 = CleanCellText(table.cell(rowIndex, 5).Range.text)
        comment2 = CleanCellText(table.cell(rowIndex, 6).Range.text)
        comment3 = CleanCellText(table.cell(rowIndex, 7).Range.text)
        On Error GoTo 0

        ' Concatenate the contents into the required format
        labelText = volume & " " & temperature & vbCrLf & _
                    group & " " & comment1 & vbCrLf & _
                    comment2 & " " & comment3

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
    ApplyColorsToLabels labelDoc, activeColor, reagentColor, corrosiveColor

    ' Save the label document
    labelDoc.SaveAs2 "CustomLabels.docx"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Sub SetupLabelTemplate(doc As Document, numRows As Integer, numColumns As Integer, rowHeight As Single)
    doc.PageSetup.PageWidth = InchesToPoints(8.5)
    doc.PageSetup.PageHeight = InchesToPoints(11)
    doc.PageSetup.TopMargin = InchesToPoints(0.5)
    doc.PageSetup.LeftMargin = InchesToPoints(0.25)
    doc.PageSetup.BottomMargin = InchesToPoints(0.5)
    doc.PageSetup.RightMargin = InchesToPoints(0.25)
    doc.PageSetup.Orientation = wdOrientPortrait

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
    End With
End Sub

Function CleanCellText(text As String) As String
    CleanCellText = Left(text, Len(text) - 2)
End Function

Function IsValidColor(color As String) As Boolean
    Select Case LCase(color)
        Case "red", "blue", "green", "yellow", "black", "white", "cyan", "magenta"
            IsValidColor = True
        Case Else
            IsValidColor = False
    End Select
End Function

Sub ApplyColorsToLabels(labelDoc As Document, activeColor As String, reagentColor As String, corrosiveColor As String)
    Dim labelRange As Range
    Dim para As Paragraph

    For Each para In labelDoc.Paragraphs
        Set labelRange = para.Range
        ApplyColorToWord labelRange, "Active", activeColor
        ApplyColorToWord labelRange, "Reagent", reagentColor
        ApplyColorToWord labelRange, "Corrosive", corrosiveColor
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