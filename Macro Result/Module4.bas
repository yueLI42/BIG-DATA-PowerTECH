Attribute VB_Name = "Module4"

Sub CopySheetToNewWorkbook()
    Dim SourceWorkbook As Workbook
    Dim NewWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    
    ' Create a new workbook
    Set NewWorkbook = Workbooks.Add

    
    row_final = ThisWorkbook.Sheets("Result").Range("C9").End(xlDown).Row
    ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set wsSource = ThisWorkbook.Sheets("Result")
    Set wsNew = NewWorkbook.Sheets.Add(Before:=NewWorkbook.Sheets(1))
        wsNew.Name = "Result"
        wsNew.Tab.Color = RGB(146, 208, 80)
    ' Copy the source worksheet to the new workbook
    wsSource.Range("A8:Z" & row_final).copy Destination:=wsNew.Range("A1")
    ' Autofit the columns in the new sheet
    wsNew.Cells.EntireColumn.AutoFit
    
    'delete useless column
     NewWorkbook.Sheets("Result").Columns("H").Delete
     NewWorkbook.Sheets("Result").Columns("J").Delete
     NewWorkbook.Sheets("Result").Range("K:N").Delete Shift:=xlToLeft
     NewWorkbook.Sheets("Result").Range("L:M").Delete Shift:=xlToLeft
     NewWorkbook.Sheets("Result").Columns("N").Delete
     NewWorkbook.Sheets("Result").Columns("P").Delete
    
    
    ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("Trainer_information")
    ' Copy the source worksheet to the new workbook
    SourceWorksheet.copy Before:=NewWorkbook.Sheets("Sheet1")
     
     'delete useless column
     NewWorkbook.Sheets("Trainer_information").Columns("F").copy
     NewWorkbook.Sheets("Trainer_information").Columns("F").PasteSpecial Paste:=xlPasteValues
     Application.CutCopyMode = False ' Clear clipboard
     NewWorkbook.Sheets("Trainer_information").Columns("E").Delete
     NewWorkbook.Sheets("Trainer_information").Columns("G").Delete
      
    ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("CAP50")
    ' Copy the source worksheet to the new workbook
    SourceWorksheet.copy Before:=NewWorkbook.Sheets("Sheet1")
    
    'delete useless column
     NewWorkbook.Sheets("CAP50").Range("C:D").Delete Shift:=xlToLeft
     NewWorkbook.Sheets("CAP50").Range("D:H").Delete Shift:=xlToLeft
     NewWorkbook.Sheets("CAP50").Range("G:J").Delete Shift:=xlToLeft
     NewWorkbook.Sheets("CAP50").Range("H:K").Delete Shift:=xlToLeft
     NewWorkbook.Sheets("CAP50").Range("I:M").Delete Shift:=xlToLeft
     

     
     Application.DisplayAlerts = False
    NewWorkbook.Sheets("Sheet1").Delete
     Application.DisplayAlerts = True
     
    ' Define the file path where you want to save the workbook
    FilePath = "C:\Users\yli6\Downloads\file_to_update\Training Result&Trainer information&CAP 50.xlsx"

    ' Create a FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Check if the file already exists
    If FSO.FileExists(FilePath) Then
        ' If the file exists, delete it first
        FSO.DeleteFile FilePath
    End If
    
    ' Save the new workbook with a desired name and path
    NewWorkbook.SaveAs "C:\Users\yli6\Downloads\file_to_update\Training Result&Trainer information&CAP 50.xlsx" ' Update with your desired file path and name
    
    ' Close the new workbook
    NewWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set NewWorkbook = Nothing
    
    
    ' Create a new workbook
    Set NewWorkbook = Workbooks.Add
    
      ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("Num_needs")
    Set NewWorksheet = NewWorkbook.Sheets.Add(Before:=NewWorkbook.Sheets("Sheet1"))
    NewWorksheet.Name = "Num_needs"
    NewWorksheet.Tab.Color = RGB(146, 208, 80)
    ' Copy the data from the source worksheet to the new worksheet
    SourceWorksheet.UsedRange.copy
    
    ' Paste the copied data as values (text) into the new worksheet
    NewWorksheet.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
      ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("Follow up list")
    ' Copy the source worksheet to the new workbook
    SourceWorksheet.copy Before:=NewWorkbook.Sheets("Sheet1")
    
      ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("LM_filter")
    ' Copy the source worksheet to the new workbook
    SourceWorksheet.copy Before:=NewWorkbook.Sheets("Sheet1")
    
    ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("Session_Completed")
    ' Copy the source worksheet to the new workbook
    SourceWorksheet.copy Before:=NewWorkbook.Sheets("Sheet1")
    
      ' Set the source worksheet (change Sheet to the name of the sheet you want to copy)
    Set SourceWorksheet = ThisWorkbook.Sheets("Training_parts_details")
    
    ' Copy the source worksheet to the new workbook
    SourceWorksheet.copy Before:=NewWorkbook.Sheets("Sheet1")
     Application.DisplayAlerts = False
    NewWorkbook.Sheets("Sheet1").Delete
     Application.DisplayAlerts = True
     
    ' Define the file path where you want to save the workbook
    FilePath = "C:\Users\yli6\Downloads\file_to_update\List of trainees & Session history.xlsx"

    ' Create a FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Check if the file already exists
    If FSO.FileExists(FilePath) Then
        ' If the file exists, delete it first
        FSO.DeleteFile FilePath
    End If
    
    ' Save the new workbook with a desired name and path
    NewWorkbook.SaveAs "C:\Users\yli6\Downloads\file_to_update\List of trainees & Session history.xlsx" ' Update with your desired file path and name
    
    ' Close the new workbook
    NewWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set NewWorkbook = Nothing
    
End Sub

