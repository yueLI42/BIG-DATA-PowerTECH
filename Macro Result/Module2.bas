Attribute VB_Name = "Module2"
Sub CopyRangeToNewSheet()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim row_final As Long
    
    ' Define the source worksheet (change "Result" to your actual sheet name)
    Set wsSource = ThisWorkbook.Sheets("Result")
    
    ' Define the last row in your source range (change as needed)
    
     row_final = Sheets("Result").Range("C9").End(xlDown).Row
    
    ' Add a new worksheet named "New Result" or use an existing one
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets("New Result")
    On Error GoTo 0
    
    If wsNew Is Nothing Then
        ' Create a new worksheet if "New Result" doesn't exist
        Set wsNew = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("Result"))
        wsNew.Name = "New Result"
        wsNew.Tab.Color = RGB(146, 208, 80)
    End If
    
    ' Copy the specified range to the new sheet
    wsSource.Range("A1:Z" & row_final).copy Destination:=wsNew.Range("A1")
    
    ' Autofit the columns in the new sheet
    wsNew.Cells.EntireColumn.AutoFit
    ' Delete the old "Result" sheet
    Application.DisplayAlerts = False ' Suppress the delete confirmation dialog
    wsSource.Delete
    Application.DisplayAlerts = True ' Re-enable the confirmation dialog
    wsNew.Name = "Result"
End Sub

