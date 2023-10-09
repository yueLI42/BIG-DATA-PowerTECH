Attribute VB_Name = "Module3"
Sub UpdateTrainerInformation()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[V5000]_IP05-6B-1___Preferred_Instructors.xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Sheet1" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("V5000 IP05-6B-1 | Preferred (2)")
    
    ' Set the destination worksheet (change "Trainer_information_source" to your destination sheet name)
    Set DestinationWorksheet = ThisWorkbook.Sheets("Trainer_information_source")
    
    ' Clear the old data in the destination worksheet
    DestinationWorksheet.Cells.clear
    
    ' Copy data from source to destination
    SourceWorksheet.UsedRange.copy DestinationWorksheet.Range("A2") ' You can change the destination range as needed
    
    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
End Sub
Sub UpdateLearningmanagement()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[Learning]_2023_Learning_Management_(In_Progress__Not_Started__Others).xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Sheet1" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("Learning 2023 Learning Mana (2)")
    
    ' Set the destination worksheet (change "Trainer_information_source" to your destination sheet name)
    Set DestinationWorksheet = ThisWorkbook.Sheets("Learning management")
    
    ' Clear the old data in the destination worksheet
    DestinationWorksheet.Cells.clear
    
    ' Copy data from source to destination
    SourceWorksheet.UsedRange.copy DestinationWorksheet.Range("A2") ' You can change the destination range as needed
    
    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
End Sub
Sub UpdateCAP()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[V5000]_PD01_6a___CAP_50_-_Overall_follow_up.xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Sheet1" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("V5000 PD01 6a | CAP 50 - Ov (2)")
    
    ' Set the destination worksheet (change "Trainer_information_source" to your destination sheet name)
    Set DestinationWorksheet = ThisWorkbook.Sheets("CAP50_follow_up_source")
    
    ' Clear the old data in the destination worksheet
    DestinationWorksheet.Cells.clear
    
    ' Copy data from source to destination
    SourceWorksheet.UsedRange.copy DestinationWorksheet.Range("A2") ' You can change the destination range as needed
    
    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
End Sub
Sub MoveLearningCompletionSheet()
    Dim SourceWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorkbook As Workbook
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[Learning]_2023_Learning_Completions.xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Learning 2023 Learning Comp (2)" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("Learning 2023 Learning Comp (2)")
    
    ' Set the destination workbook (change "YourDestinationWorkbookName.xlsx" to your destination workbook name)
    Set DestinationWorkbook = ThisWorkbook
    
     
    ' Check if the destination worksheet already exists in the destination workbook
    On Error Resume Next
    Dim DestSheet As Worksheet
    Set DestSheet = DestinationWorkbook.Sheets("Learning completion")
    On Error GoTo 0
    
    ' If the destination worksheet exists, delete it first
    If Not DestSheet Is Nothing Then
        Application.DisplayAlerts = False ' Turn off alerts to delete the sheet
        DestSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Move the source worksheet to the destination workbook
    SourceWorksheet.Move Before:=DestinationWorkbook.Sheets(12)
    ' Set the tab color of the moved worksheet (change RGB values for the desired color)
    DestinationWorkbook.Sheets(12).Tab.Color = RGB(255, 192, 0)
    DestinationWorkbook.Sheets(12).Name = "Learning completion"
    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
    Set DestinationWorkbook = Nothing
End Sub

Sub UpdateAllMyltrainer()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[USERS]_All_MyLearning_Trainers_from_my_perimeter.xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Sheet1" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("USERS All MyLearning Trainers f")
    
    ' Set the destination worksheet (change "Trainer_information_source" to your destination sheet name)
    Set DestinationWorksheet = ThisWorkbook.Sheets("All_Myl_trainer")
    
    ' Clear the old data in the destination worksheet
    DestinationWorksheet.Cells.clear
    
    ' Copy data from source to destination
    SourceWorksheet.UsedRange.copy DestinationWorksheet.Range("A2") ' You can change the destination range as needed
    
    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
End Sub
Sub UpdateSessionsfollow()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[KPIs]_2023_Training_sessions_follow-up_(Assigned_in_2023).xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Sheet1" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("KPIs 2023 Training sessions fol")
    
    ' Set the destination worksheet (change "Trainer_information_source" to your destination sheet name)
    Set DestinationWorksheet = ThisWorkbook.Sheets("Sessions follow up source")
    
    ' Clear the old data in the destination worksheet
    DestinationWorksheet.Cells.clear
    
    ' Copy data from source to destination
    SourceWorksheet.UsedRange.copy DestinationWorksheet.Range("A1") ' You can change the destination range as needed
    
    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
End Sub
Sub UpdateCatalog()
    Dim SourceWorkbook As Workbook
    Dim DestinationWorkbook As Workbook
    Dim SourceWorksheet As Worksheet
    Dim DestinationWorksheet As Worksheet
    
    ' Define the source file path
    Dim SourceFilePath As String
    SourceFilePath = "C:\Users\yli6\Downloads\source\[CATALOG]_!_Full_MyLearning_Catalog_!.xlsx" ' Update with your source file path
    
    ' Check if the source file exists
    If Dir(SourceFilePath) = "" Then
        MsgBox "Source file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Open the source workbook
    Set SourceWorkbook = Workbooks.Open(SourceFilePath)
    
    ' Set the source worksheet (change "Sheet1" to the name of your source sheet)
    Set SourceWorksheet = SourceWorkbook.Sheets("CATALOG ! Full MyLearning Catal")
    
    ' Set the destination worksheet (change "Trainer_information_source" to your destination sheet name)
    Set DestinationWorksheet = ThisWorkbook.Sheets("Catalog")
    
    ' Clear the old data in the destination worksheet
    DestinationWorksheet.Cells.clear
    
    row1 = SourceWorksheet.Range("A13").End(xlDown).Row
    
    SourceWorksheet.Range("A13", "AA" & row1).AutoFilter Field:=4, Criteria1:=Array _
        ("Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"), Operator:= _
        xlFilterValues
    
    ' Copy data from source to destination
    row_final = SourceWorksheet.Range("A13").End(xlDown).Row
    SourceWorksheet.Range("A13:A" & row_final).copy Destination:=DestinationWorksheet.Range("A1")
    SourceWorksheet.Range("D13:D" & row_final).copy Destination:=DestinationWorksheet.Range("B1")
    SourceWorksheet.Range("F13:F" & row_final).copy Destination:=DestinationWorksheet.Range("C1")
    SourceWorksheet.Range("U13:U" & row_final).copy Destination:=DestinationWorksheet.Range("D1")
    'delete duplicate
    DestinationWorksheet.Range("A1", "D" & row_final).RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes

    ' Close the source workbook without saving
    SourceWorkbook.Close SaveChanges:=False
    
    ' Clean up
    Set SourceWorksheet = Nothing
    Set DestinationWorksheet = Nothing
    Set SourceWorkbook = Nothing
End Sub


