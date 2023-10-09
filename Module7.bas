Attribute VB_Name = "Module7"
Sub Session_Completed()

     Sheets("Sessions follow up source").Rows("1:1").AutoFilter
    On Error Resume Next
        Sheets("Sessions follow up source").AutoFilter.Sort.SortFields.clear
    If Err <> 0 Then
        Sheets("Sessions follow up source").Rows("1:1").AutoFilter
    End If
    Err.clear
    
    row2 = Sheets("Sessions follow up source").Range("A1").End(xlDown).Row
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=11, Criteria1:= _
        "Session"
    
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=28, Criteria1:= _
        "Completed"
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=29, Criteria1:= _
        "Completed"
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=32, Criteria1:=Array( _
        "Central R&D", "Group R&D", "PowerTECH Knowledge"), Operator:= _
        xlFilterValues
    
    row_initial = Sheets("Session_Completed").Range("A1").End(xlDown).Row
    
    Sheets("Session_Completed").Range("A1", "C" & row_initial).ClearContents
    
    Sheets("Sessions follow up source").Range("A1", "A" & row2).copy Destination:=Sheets("Session_Completed").Range("A1")
    Sheets("Sessions follow up source").Range("N1", "N" & row2).copy Destination:=Sheets("Session_Completed").Range("B1")
    Sheets("Sessions follow up source").Range("W1", "W" & row2).copy Destination:=Sheets("Session_Completed").Range("C1")
    'delete duplicate
    Sheets("Session_Completed").Range("A1", "C" & row2).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    Sheets("Session_Completed").Columns("B").Delete
    Row_total = Sheets("Session_Completed").Range("A1").End(xlDown).Row
    Sheets("Session_Completed").Rows("1:1").AutoFilter ' filter reset
    Sheets("Session_Completed").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B1", "B" & Row_total), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With Sheets("Session_Completed").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Count_PTS = 0
    Count_No_PTS = 0
    Line = 2
    For i = 2 To Row_total
        If Sheets("Session_Completed").Cells(i, 2) <> Sheets("Session_Completed").Cells(i + 1, 2) Then
                 If Sheets("Session_Completed").Cells(i, 1) = "PTS Powertrain Systems" Then
                     Count_PTS = Count_PTS + 1
                Else
                     Count_No_PTS = Count_No_PTS + 1
                End If
              Sheets("Session_Completed").Cells(Line, 3).Value = Sheets("Session_Completed").Cells(i, 2)
              Sheets("Session_Completed").Cells(Line, 4).Value = Count_PTS
              Sheets("Session_Completed").Cells(Line, 5).Value = Count_No_PTS
              Line = Line + 1
              Count_PTS = 0
              Count_No_PTS = 0
              
        Else
            If Sheets("Session_Completed").Cells(i, 1) = "PTS Powertrain Systems" Then
                 Count_PTS = Count_PTS + 1
            Else
                 Count_No_PTS = Count_No_PTS + 1
            End If
            
        End If
    Next
    
    Sheets("Session_Completed").Columns("A").Delete
    Sheets("Session_Completed").Columns("A").Delete
    Sheets("Session_Completed").Cells(1, 1).Value = "Locator Number"
    Sheets("Session_Completed").Cells(1, 2).Value = "Count_PTS"
    Sheets("Session_Completed").Cells(1, 3).Value = "Count_No_PTS"
    
    Sheets("Sessions follow up source").Rows("1:1").AutoFilter ' filter reset
    On Error Resume Next
    Sheets("Session_Completed").AutoFilterMode = False
    On Error GoTo 0
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim row_final As Long
    
    ' Define the source worksheet (change "Result" to your actual sheet name)
    Set wsSource = ThisWorkbook.Sheets("Session_Completed")
    
    ' Define the last row in your source range (change as needed)
    
     row_final = Sheets("Session_Completed").Range("C1").End(xlDown).Row
    
    ' Add a new worksheet named "New Result" or use an existing one
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets("New Session_Completed")
    On Error GoTo 0
    
    If wsNew Is Nothing Then
        ' Create a new worksheet if "New Result" doesn't exist
        Set wsNew = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("Catalog"))
        wsNew.Name = "New Result"
        wsNew.Tab.Color = RGB(146, 208, 80)
    End If
    
    ' Copy the specified range to the new sheet
    wsSource.Range("A1:C" & row_final).copy Destination:=wsNew.Range("A1")
    
    ' Autofit the columns in the new sheet
    wsNew.Cells.EntireColumn.AutoFit
    ' Delete the old "Result" sheet
    Application.DisplayAlerts = False ' Suppress the delete confirmation dialog
    wsSource.Delete
    Application.DisplayAlerts = True ' Re-enable the confirmation dialog
    wsNew.Name = "Session_Completed"
    
End Sub



