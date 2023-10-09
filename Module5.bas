Attribute VB_Name = "Module5"
Sub learning_completions_old()

  Application.StatusBar = "test"
        
    row_com = Sheets("Learning completion old").Range("A1").End(xlDown).Row
    row_result = Sheets("Result").Range("A9").End(xlDown).Row
           
        
        
    Sheets("Learning completion old").Range("A2", "Y" & row_com).copy Destination:=Sheets("Result").Range("A" & row_result + 1)
    
    
    'row_result2 = Sheets("Result").Range("A9").End(xlDown).Row
   
    'Sheets("Result").Range("A" & row_result + 1, "Y" & row_result2).Interior.Color = RGB(173, 216, 0)
    
    
    ' Specify the column containing the dates (column A in this example)
    'Set dateColumn = Sheets("Result").Range("J" & row_result + 1 & ":J" & row_result2)
    
    ' Change the format of the date column to the desired date format
    'dateColumn.NumberFormat = "mm/dd/yyyy hh:mm AM/PM" ' Change the format as needed
    
    ' Optionally, convert the text to actual dates if needed
    'dateColumn.Value = dateColumn.Value
    
  
   
End Sub
