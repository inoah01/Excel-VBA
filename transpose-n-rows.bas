Attribute VB_Name = "Module1"
Sub Transpose_N_Rows()
' This is a macro allows N rows of data to be transposed onto another sheet.

' Make the macro run faster on large data sets.
Application.ScreenUpdating = False

' Collecting user-selected rows and columns.
xRow = Selection.Rows.Count
xCol = Selection.Column

' The row that the Transposed data will be put into.
nextRow = InputBox("How many rows should the ouput be offset by?")

' How many rows to Transpose.
stepValue = InputBox("How many rows should be grouped together?")

' Destination for the transposed data.
Destination = InputBox("What is the name of the destination sheet?")

' Loop through the user-selected data using a step value.
For i = 1 To xRow Step stepValue
    
    ' Copy the data, using the step value to determine the size of
    ' the copied range.
    Cells(i, xCol).Resize(stepValue).Copy
    
    'Transpose the data.
    Sheets(Destination).Cells(1, xCol).Offset(nextRow, 0).PasteSpecial Paste:=xlPasteAll, Transpose:=True
    
    ' Increment the nextRow value so the copied data goes onto
    ' a new line.
    nextRow = nextRow + 1

Next

' Remove the "copy lines" from the Transposed data.
Application.CutCopyMode = False
' Make Excel function as expected after the macro is finished.
Application.ScreenUpdating = True

End Sub

