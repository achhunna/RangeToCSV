Sub csv()

    Dim Path As String
    Dim Filename As String
    Dim ActiveRange As Range
    Dim RangeValues() As Variant
    Dim OutputFile As Long

    On Error GoTo ErrHndlr
    Filename = Range("A1").Value
    Path = Application.ActiveWorkbook.Path

    ReDim RangeValues(1 To Selection.Rows.Count)
    For i = 1 To Selection.Rows.Count
         On Error GoTo ErrHndlr
         RangeValues(i) = Join(WorksheetFunction.Transpose(WorksheetFunction. _
         Transpose(Selection.Rows(i).Value)), ",")
    Next

    OutputFile = FreeFile

    Open Path & "\" & Filename & ".csv" For Output Lock Write As #OutputFile

    Print #OutputFile, Join(RangeValues, vbNewLine)
    Close OutputFile
ErrHndlr:

End Sub
