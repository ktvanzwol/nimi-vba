Attribute VB_Name = "niTools_ExamplesSupport"
Option Explicit


Function Example_GetNewOutputSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim sh As Shape
    
    On Error GoTo Error:
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
'    Application.DisplayAlerts = False
'    ws.Delete
'    Application.DisplayAlerts = True

    ws.UsedRange.Clear
    
    For Each sh In ws.Shapes
        sh.Delete
    Next

    
    Set Example_GetNewOutputSheet = ws
    Exit Function
Error:
    
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
    ws.name = sheetName
    
    Set Example_GetNewOutputSheet = ws
End Function



Public Sub ClearRows(startCell As range, columns As Long)
    Dim r As Long
    Dim c As Long
    Dim i As Long
      
    r = startCell.Row
    c = startCell.Column
    
    Do While Cells(r, c).Value2 <> ""
        For i = 0 To columns - 1
            Cells(r, c + i).Value2 = ""
        Next
        r = r + 1
    Loop
    
End Sub

