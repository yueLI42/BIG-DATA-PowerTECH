Attribute VB_Name = "Module1"
Sub trainer_information_copy()
    'clear the trainer information
    row1 = Sheets("Trainer_information").Range("A1").End(xlDown).Row
    Sheets("Trainer_information").Range("A1", "J" & row1).ClearContents
    With Sheets("Trainer_information").Range("A1", "J" & row1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Sheets("Trainer_information_source").Rows("12:12").AutoFilter
    On Error Resume Next
        Sheets("Trainer_information_source").AutoFilter.Sort.SortFields.clear
    If Err <> 0 Then
        Sheets("Trainer_information_source").Rows("12:12").AutoFilter
    End If
    Err.clear

    row_trainer = Sheets("Trainer_information_source").Range("B12").End(xlDown).Row
        
    Sheets("Trainer_information_source").Range("A12", "H" & row_trainer).AutoFilter Field:=8, Criteria1:="<>"
        
    row_trainer2 = Sheets("Trainer_information_source").Range("B12").End(xlDown).Row
    'insert the site & country
    Sheets("Trainer_information_source").Range("B12", "H" & row_trainer2).copy Destination:=Sheets("Trainer_information").Range("A1")
    
    row2 = Sheets("Trainer_information").Range("A1").End(xlDown).Row
    
    Sheets("Trainer_information").Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Trainer_information").Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Sheets("Trainer_information").Range("F1").FormulaR1C1 = "Site"
    Sheets("Trainer_information").Range("G1").FormulaR1C1 = "Country"
    
    Sheets("Trainer_information").Range("F2").FormulaR1C1 = "=IF(ISERR(FIND(""- "",RC[-1])),RC[-1],RIGHT(RC[-1],LEN(RC[-1])-FIND(""- "",RC[-1])-1))"
    Sheets("Trainer_information").Range("F2").AutoFill Destination:=Sheets("Trainer_information").Range("F2", "F" & row2)
    
    
    'store training title
    'Dim training_title As Object
    'Set training_title = CreateObject("scripting.dictionary")
    'row_title = Sheets("CVM_Title & Type").Range("A1").End(xlDown).Row
    'For i = 2 To row_title
     '   training_title.Add Sheets("CVM_Title & Type").Range("A" & i).Value, Sheets("CVM_Title & Type").Range("B" & i).Value
    'Next
    
    'insert the country
    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    row_site = Sheets("Site_country").Range("A1").End(xlDown).Row
    For i = 2 To row_site
        d.Add Sheets("Site_country").Range("A" & i).Value, Sheets("Site_country").Range("B" & i).Value
    Next
    
    Dim lastRowSite As Long
    lastRowSite = Sheets("Site_country").Cells(Sheets("Site_country").Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To row2
        'insert the country
        Dim SiteName As String
       SiteName = Sheets("Trainer_information").Range("F" & i).Value
    
      If Len(SiteName) > 0 Then ' Check if SiteName is not empty
        If d.Exists(SiteName) Then
            Sheets("Trainer_information").Range("G" & i).Value = d(SiteName)
        Else
            ' If there's no match, put the value in the last row of Site_country
            lastRowSite = lastRowSite + 1
            Sheets("Site_country").Range("A" & lastRowSite).Value = SiteName
            Sheets("Site_country").Range("B" & lastRowSite).Value = ""
        End If
     End If
        'reset the title
        'title = Sheets("Trainer_information").Range("A" & i).Value
        'If (training_title(title) = "") Then
         '   Sheets("Trainer_information").Range("A" & i).Value = title
       ' Else
        '    Sheets("Trainer_information").Range("A" & i).Value = training_title(title)
        'End If
    Next
    
    ' Remove duplicates in the Site_country worksheet
    Sheets("Site_country").Range("A1:B" & lastRowSite).RemoveDuplicates Columns:=Array(1), Header:=xlNo

    
    ' source reset
    Sheets("Trainer_information_source").Rows("12:12").AutoFilter

    
    Call add_BG  'add the information "BG"
    
End Sub

Sub cvv_copy()
    Sheets("CVV").Rows("2:2").AutoFilter
    On Error Resume Next
        Sheets("CVV").AutoFilter.Sort.SortFields.clear
        
    row_cvv = Sheets("CVV").Range("A2").End(xlDown).Row
    
    If Err <> 0 Then
        Sheets("CVV").Rows("2:2").AutoFilter
    End If
    Err.clear
    
    Sheets("CVV").Columns("AF:AF").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("CVV").Range("AF1").FormulaR1C1 = "Status"
    Sheets("CVV").Range("AF2", "AF" & row_cvv).FormulaR1C1 = "=IF(RC[-1]=""100.00%"",""Completed"",""Not Started"")"
   
   
    Sheets("CVV").Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("CVV").Range("I1").FormulaR1C1 = "Site"
    Sheets("CVV").Range("I2", "I" & row_cvv).FormulaR1C1 = "=RIGHT(RC[-1],4)"
   
   
    Sheets("CVV").Range("A2", "AP" & row_cvv).AutoFilter Field:=18, Criteria1:=Array _
        ("Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"), Operator:= _
        xlFilterValues
        
    row_result = Sheets("Result").Range("C8").End(xlDown).Row
        
    
    Sheets("CVV").Range("B3", "B" & row_cvv).copy
    Sheets("Result").Range("C" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Sheets("CVV").Range("N3", "N" & row_cvv).copy
    Sheets("Result").Range("D" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Sheets("CVV").Range("F3", "F" & row_cvv).copy
    Sheets("Result").Range("E" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Sheets("CVV").Range("J3", "J" & row_cvv).copy  'title
    Sheets("Result").Range("G" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("CVV").Range("S3", "S" & row_cvv).copy  'type
    Sheets("Result").Range("H" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("CVV").Range("T3", "T" & row_cvv).copy 'position id
    Sheets("Result").Range("W" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("CVV").Range("K3", "K" & row_cvv).copy  'training id
    Sheets("Result").Range("Y" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("CVV").Range("D3", "D" & row_cvv).copy 'manager id
    Sheets("Result").Range("N" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("CVV").Range("R3", "R" & row_cvv).copy 'provider
    Sheets("Result").Range("Q" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    row_result2 = Sheets("Result").Range("C10").End(xlDown).Row
    Sheets("Result").Range("S" & row_result + 1, "S" & row_result2).FormulaR1C1 = "CVM"
    'Sheets("Result").Range("S" & row_result + 1).AutoFill Destination:=Sheets("Result").Range("L" & row_result + 1, "L" & row_result2)
    
    Sheets("CVV").Range("I3", "I" & row_cvv).copy 'site
    Sheets("Result").Range("F" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("Result").Range("L" & row_result + 1, "L" & row_result2).Value = Sheets("CVV").Range("AG3", "AG" & row_cvv).Value
' data source reset

    Dim arr1()
    row1 = Sheets("Result").Range("C" & row_result + 1).End(xlDown).Row
    arr1 = Sheets("Result").Range("C" & row_result + 1, "Y" & row_result2).Value
    
    For i = 1 To row_result2 - row_result
        If arr1(i, 6) = "Online Course" Then
            Sheets("Result").Range("H" & i + row_result).Value = "Online Class" ' reset online course
        End If
        If arr1(i, 2) = "PEM" Then
            Sheets("Result").Range("D" & i + row_result).Value = "PEM Powertrain  Electrified Mobility" 'reset pem
        ElseIf arr1(i, 2) = "PSD" Then
            Sheets("Result").Range("D" & i + row_result).Value = "PSD Powertrain Systems Driveline" 'reset psd
        ElseIf arr1(i, 2) = "Empty" Then
            Sheets("Result").Range("D" & i + row_result).Value = ""
        End If
    Next
    
    row_cvv = Sheets("CVV").Range("A2").End(xlDown).Row
    Sheets("CVV").Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("CVV").Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("CVV").Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("CVV").Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Sheets("CVV").Range("C3", "C" & row_cvv).FormulaR1C1 = "=LEFT(RC[-1],FIND(""@"",RC[-1])-1)"
    Sheets("CVV").Range("D3", "D" & row_cvv).FormulaR1C1 = "=LEFT(RC[-1],FIND(""."",RC[-1])-1)"
    Sheets("CVV").Range("E3", "E" & row_cvv).FormulaR1C1 = "=RIGHT(RC[-2],LEN(RC[-2])-FIND(""."",RC[-2]))"
    Sheets("CVV").Range("F3", "F" & row_cvv).FormulaR1C1 = "=IF(ISNUMBER(FIND(""."",RC[-1])),LEFT(RC[-1],FIND(""."",RC[-1])-1),RC[-1])"


    For i = 3 To row_cvv

        Sheets("CVV").Range("D" & i).Value = StrConv(Sheets("CVV").Range("D" & i).Value, vbProperCase)
        Sheets("CVV").Range("F" & i).Value = StrConv(Sheets("CVV").Range("F" & i).Value, vbUpperCase)

    Next
    

    Sheets("CVV").Columns("E:E").Delete Shift:=xlToLeft
    Sheets("CVV").Columns("C:C").Delete Shift:=xlToLeft
    Sheets("CVV").Range("C2").FormulaR1C1 = "First name"
    Sheets("CVV").Range("D2").FormulaR1C1 = "Last name"

    Sheets("CVV").Range("D3", "D" & row_cvv).copy 'last name
    Sheets("Result").Range("A" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Sheets("CVV").Range("C3", "C" & row_cvv).copy 'first name
    Sheets("Result").Range("B" & row_result + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Define the range of the first column
    Set FirstColumn = Sheets("Result").Range("A" & 9 & ":A" & row_result2 - row_result + 8) ' Change "A:A" to the column you want to copy from
    ' Define the range of the other columns where you want to apply the formatting
    Set OtherColumns = Sheets("Result").Range("A" & row_result + 1 & ":A" & row_result2) ' Change "R:Z" to the range of columns you want to format
    ' Copy the formatting of the first column and apply it to the other columns
    FirstColumn.copy
    OtherColumns.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear the clipboard
    Sheets("Result").Range("A" & row_result + 1, "A" & row_result2).Interior.Color = RGB(255, 200, 100)

    ' Define the range of the first column
    Set FirstColumn = Sheets("Result").Range("A" & row_result + 1 & ":A" & row_result2) ' Change "A:A" to the column you want to copy from
    ' Define the range of the other columns where you want to apply the formatting
    Set OtherColumns = Sheets("Result").Range("B" & row_result + 1 & ":Y" & row_result2) ' Change "R:Z" to the range of columns you want to format
     
    ' Copy the formatting of the first column and apply it to the other columns
    FirstColumn.copy
    OtherColumns.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear the clipboard
    
    Sheets("CVV").Rows("1:1").AutoFilter
    If Sheets("CVV").Range("D2").Value = "Last name" Then
        Sheets("CVV").Columns("D:D").Delete Shift:=xlToLeft
    End If
    If Sheets("CVV").Range("C2").Value = "First name" Then
        Sheets("CVV").Columns("C:C").Delete Shift:=xlToLeft
    End If
    If Sheets("CVV").Range("AG1").Value = "Status" Then
        Sheets("CVV").Columns("AG:AG").Delete Shift:=xlToLeft
    End If
    If Sheets("CVV").Range("I1").Value = "Site" Then
        Sheets("CVV").Columns("I:I").Delete Shift:=xlToLeft
    End If
    

End Sub

Sub learning_management_copy()
    
    Sheets("Learning management").Rows("9:9").AutoFilter
    On Error Resume Next
        Sheets("Learning management").AutoFilter.Sort.SortFields.clear
        'Sheets("Learning management").Rows("9:9").AutoFilter
    

    row2 = Sheets("Learning management").Range("A9").End(xlDown).Row
    
    
    Sheets("Learning management").Columns("M:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Learning management").Range("M9").FormulaR1C1 = "Site"
    Sheets("Learning management").Range("M10").FormulaR1C1 = "=IF(ISERR(FIND(""- "",RC[-1])),RC[-1],RIGHT(RC[-1],LEN(RC[-1])-FIND(""- "",RC[-1])-1))"
    Sheets("Learning management").Range("M10").AutoFill Destination:=Sheets("Learning management").Range("M10", "M" & row2)
    ' Convert formulas to values in the entire column M
  Sheets("Learning management").Range("M10", "M" & row2).Value = Sheets("Learning management").Range("M10", "M" & row2).Value
    '************
    
    'attention insert one line, the number of the line has changed
    
    If Err <> 0 Then
        Sheets("Learning management").Rows("9:9").AutoFilter
    End If
    Err.clear
    
    
    Sheets("Learning management").Range("A9", "AO" & row2).AutoFilter Field:=6, Criteria1:= _
        "PTS Powertrain Systems"

    Sheets("Learning management").Range("A9", "AO" & row2).AutoFilter Field:=30, Criteria1:=Array _
        ("Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"), Operator:= _
        xlFilterValues
    
    Sheets("Learning management").AutoFilter.Sort.SortFields. _
        clear
    Sheets("Learning management").AutoFilter.Sort.SortFields. _
        Add Key:=Range("D9", "D" & row2), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With Sheets("Learning management").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("Learning management").Range("B10", "B" & row2).copy Destination:=Sheets("Result").Range("A9")
    Sheets("Learning management").Range("C10", "C" & row2).copy Destination:=Sheets("Result").Range("B9")
    Sheets("Learning management").Range("D10", "D" & row2).copy Destination:=Sheets("Result").Range("C9")
    Sheets("Learning management").Range("G10", "G" & row2).copy Destination:=Sheets("Result").Range("D9")
    Sheets("Learning management").Range("K10", "K" & row2).copy Destination:=Sheets("Result").Range("E9")
    Sheets("Learning management").Range("M10", "M" & row2).copy Destination:=Sheets("Result").Range("F9")
    Sheets("Learning management").Range("R10", "R" & row2).copy Destination:=Sheets("Result").Range("G9")
    Sheets("Learning management").Range("S10", "S" & row2).copy Destination:=Sheets("Result").Range("H9")
    Sheets("Learning management").Range("T10", "T" & row2).copy Destination:=Sheets("Result").Range("I9")
    Sheets("Learning management").Range("V10", "V" & row2).copy Destination:=Sheets("Result").Range("K9")
    Sheets("Learning management").Range("W10", "W" & row2).copy Destination:=Sheets("Result").Range("L9")
    Sheets("Learning management").Range("X10", "X" & row2).copy Destination:=Sheets("Result").Range("M9")
    Sheets("Learning management").Range("Y10", "Y" & row2).copy Destination:=Sheets("Result").Range("N9")
    Sheets("Learning management").Range("AF10", "AF" & row2).copy Destination:=Sheets("Result").Range("O9")
    Sheets("Learning management").Range("AG10", "AG" & row2).copy Destination:=Sheets("Result").Range("P9")
    Sheets("Learning management").Range("AD10", "AD" & row2).copy Destination:=Sheets("Result").Range("Q9")
    Sheets("Learning management").Range("N10", "N" & row2).copy Destination:=Sheets("Result").Range("W9")
    row2 = Sheets("Result").Range("A9").End(xlDown).Row
    Sheets("Result").Range("A" & 9, "Y" & row2).Interior.Color = RGB(200, 162, 200)
    ' Define the range of the first column
    Set FirstColumn = Sheets("Result").Range("A9:A" & row2) ' Change "A:A" to the column you want to copy from
    
    ' Define the range of the other columns where you want to apply the formatting
    Set OtherColumns = Sheets("Result").Range("J9:Y" & row2) ' Change "R:Z" to the range of columns you want to format
    
    ' Copy the formatting of the first column and apply it to the other columns
    FirstColumn.copy
    OtherColumns.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear the clipboard
    
    ' Specify the column containing the dates (column A in this example)
    Set dateColumn = Sheets("Result").Range("O9:P" & row2)
    
    ' Change the format of the date column to the desired date format
    dateColumn.NumberFormat = "mm/dd/yyyy hh:mm AM/PM" ' Change the format as needed
    
    ' Optionally, convert the text to actual dates if needed
    dateColumn.Value = dateColumn.Value
    ' data source reset
    Sheets("Learning management").Rows("9:9").AutoFilter
    If Sheets("Learning management").Range("M9").Value = "Site" Then
        Sheets("Learning management").Columns("M:M").Delete Shift:=xlToLeft
    End If
    
    'Sheets("result").Range("M9").Value = "Target time"
    
End Sub

Sub learning_completion_copy()

    Sheets("Learning completion").Rows("9:9").AutoFilter
    On Error Resume Next
        Sheets("Learning completion").AutoFilter.Sort.SortFields.clear
        
    row_com = Sheets("Learning completion").Range("B9").End(xlDown).Row
    
    If Err <> 0 Then
        Sheets("Learning completion").Rows("9:9").AutoFilter
    End If
    Err.clear
    
    Sheets("Learning completion").Columns("O:O").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Learning completion").Range("O9").FormulaR1C1 = "Site"
    Sheets("Learning completion").Range("O14").FormulaR1C1 = "=IF(ISERR(FIND(""- "",RC[-1])),RC[-1],RIGHT(RC[-1],LEN(RC[-1])-FIND(""- "",RC[-1])-1))"
    Sheets("Learning completion").Range("O14").AutoFill Destination:=Sheets("Learning completion").Range("O14", "O" & row_com)
    ' Convert formulas to values in the entire column O
    Sheets("Learning completion").Range("O14", "O" & row_com).Value = Sheets("Learning completion").Range("O14", "O" & row_com).Value
    Sheets("Learning completion").Range("A9", "AK" & row_com).AutoFilter Field:=8, Criteria1:= _
        "PTS Powertrain Systems"
    
    Sheets("Learning completion").Range("A9", "AK" & row_com).AutoFilter Field:=23, Criteria1:= _
        "=Completed", Operator:=xlOr, Criteria2:="=Completed (Equivalent)"
   
    Sheets("Learning completion").Range("A9", "AK" & row_com).AutoFilter Field:=37, Criteria1:=Array _
        ("Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"), Operator:= _
        xlFilterValues
        
    row_result = Sheets("Result").Range("A8").End(xlDown).Row

    
        
    Sheets("Learning completion").Range("C14", "C" & row_com).copy Destination:=Sheets("Result").Range("A" & row_result + 1)
    Sheets("Learning completion").Range("D14", "D" & row_com).copy Destination:=Sheets("Result").Range("B" & row_result + 1)
    Sheets("Learning completion").Range("E14", "E" & row_com).copy Destination:=Sheets("Result").Range("C" & row_result + 1)
    Sheets("Learning completion").Range("I14", "I" & row_com).copy Destination:=Sheets("Result").Range("D" & row_result + 1)
    Sheets("Learning completion").Range("M14", "M" & row_com).copy Destination:=Sheets("Result").Range("E" & row_result + 1)
    Sheets("Learning completion").Range("O14", "O" & row_com).copy Destination:=Sheets("Result").Range("F" & row_result + 1)
    Sheets("Learning completion").Range("T14", "T" & row_com).copy Destination:=Sheets("Result").Range("G" & row_result + 1)
    Sheets("Learning completion").Range("U14", "U" & row_com).copy Destination:=Sheets("Result").Range("H" & row_result + 1)
    Sheets("Learning completion").Range("V14", "V" & row_com).copy Destination:=Sheets("Result").Range("I" & row_result + 1)
    Sheets("Learning completion").Range("X14", "X" & row_com).copy Destination:=Sheets("Result").Range("J" & row_result + 1)
    Sheets("Learning completion").Range("W14", "W" & row_com).copy Destination:=Sheets("Result").Range("K" & row_result + 1)
    Sheets("Learning completion").Range("P14", "P" & row_com).copy Destination:=Sheets("Result").Range("W" & row_result + 1)
    
    
    row_result2 = Sheets("Result").Range("A8").End(xlDown).Row

    Sheets("Result").Range("L" & row_result + 1).FormulaR1C1 = "Completed"
    Sheets("Result").Range("L" & row_result + 1).AutoFill Destination:=Sheets("Result").Range("L" & row_result + 1, "L" & row_result2)
    
    Sheets("Learning completion").Range("Y14", "Y" & row_com).copy Destination:=Sheets("Result").Range("M" & row_result + 1)
    Sheets("Learning completion").Range("AA14", "AA" & row_com).copy Destination:=Sheets("Result").Range("N" & row_result + 1)
    Sheets("Learning completion").Range("AK14", "AK" & row_com).copy Destination:=Sheets("Result").Range("Q" & row_result + 1)
    
    Sheets("Result").Range("A" & row_result + 1, "Y" & row_result2).Interior.Color = RGB(193, 237, 194)
    ' Define the range of the first column
    Set FirstColumn = Sheets("Result").Range("A" & row_result + 1 & ":A" & row_result2) ' Change "A:A" to the column you want to copy from
    
    ' Define the range of the other columns where you want to apply the formatting
    Set OtherColumns = Sheets("Result").Range("J" & row_result + 1 & ":Y" & row_result2) ' Change "R:Z" to the range of columns you want to format
    
    ' Copy the formatting of the first column and apply it to the other columns
    FirstColumn.copy
    OtherColumns.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear the clipboard
    ' Specify the column containing the dates (column A in this example)
    Set dateColumn = Sheets("Result").Range("J" & row_result + 1 & ":J" & row_result2)
    
    ' Change the format of the date column to the desired date format
    dateColumn.NumberFormat = "mm/dd/yyyy hh:mm AM/PM" ' Change the format as needed
    
    ' Optionally, convert the text to actual dates if needed
    dateColumn.Value = dateColumn.Value
    
    ' data source reset
    Sheets("Learning completion").Rows("9:9").AutoFilter
    
    If Sheets("Learning completion").Range("O9").Value = "Site" Then
        Sheets("Learning completion").Columns("O:O").Delete Shift:=xlToLeft
    End If
    'Sheets("result").Range("M9").Value = "Target time"


End Sub
Sub completion_add()

  Application.StatusBar = "test"
        
    row_com = Sheets("Learning completion 2022").Range("A1").End(xlDown).Row
    
    Sheets("Learning completion 2022").Columns("N:N").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Learning completion 2022").Range("N1").FormulaR1C1 = "Site"
    Sheets("Learning completion 2022").Range("N2", "N" & row_com).FormulaR1C1 = "=IF(ISERR(FIND(""- "",RC[-1])),RC[-1],RIGHT(RC[-1],LEN(RC[-1])-FIND(""- "",RC[-1])-1))"
    Sheets("Learning completion 2022").Range("N2").AutoFill Destination:=Sheets("Learning completion 2022").Range("N2", "N" & row_com)
    ' Convert formulas to values in the entire column N
    Sheets("Learning completion 2022").Range("N2", "N" & row_com).Value = Sheets("Learning completion 2022").Range("N2", "N" & row_com).Value
        
    row_result = Sheets("Result").Range("A8").End(xlDown).Row
        
        
        
    Sheets("Learning completion 2022").Range("B2", "B" & row_com).copy Destination:=Sheets("Result").Range("A" & row_result + 1) 'last name
    Sheets("Learning completion 2022").Range("C2", "C" & row_com).copy Destination:=Sheets("Result").Range("B" & row_result + 1) 'first name
    Sheets("Learning completion 2022").Range("D2", "D" & row_com).copy Destination:=Sheets("Result").Range("C" & row_result + 1) 'id
    Sheets("Learning completion 2022").Range("H2", "H" & row_com).copy Destination:=Sheets("Result").Range("D" & row_result + 1) 'pg
    Sheets("Learning completion 2022").Range("L2", "L" & row_com).copy Destination:=Sheets("Result").Range("E" & row_result + 1) 'country
    Sheets("Learning completion 2022").Range("N2", "N" & row_com).copy Destination:=Sheets("Result").Range("F" & row_result + 1) 'site
    Sheets("Learning completion 2022").Range("R2", "R" & row_com).copy Destination:=Sheets("Result").Range("G" & row_result + 1) 'title
    Sheets("Learning completion 2022").Range("S2", "S" & row_com).copy Destination:=Sheets("Result").Range("H" & row_result + 1) 'type
    Sheets("Learning completion 2022").Range("T2", "T" & row_com).copy Destination:=Sheets("Result").Range("I" & row_result + 1) 'assigned date
    Sheets("Learning completion 2022").Range("V2", "V" & row_com).copy Destination:=Sheets("Result").Range("J" & row_result + 1) 'completed date
    Sheets("Learning completion 2022").Range("U2", "U" & row_com).copy Destination:=Sheets("Result").Range("K" & row_result + 1) 'ts
    Sheets("Learning completion 2022").Range("O2", "O" & row_com).copy Destination:=Sheets("Result").Range("W" & row_result + 1) 'position id
    Sheets("Learning completion 2022").Range("W2", "W" & row_com).copy Destination:=Sheets("Result").Range("M" & row_result + 1) 'training hours
    Sheets("Learning completion 2022").Range("Y2", "Y" & row_com).copy Destination:=Sheets("Result").Range("N" & row_result + 1) 'manager id
    Sheets("Learning completion 2022").Range("AI2", "AI" & row_com).copy Destination:=Sheets("Result").Range("Q" & row_result + 1) 'provider
    
    row_result2 = Sheets("Result").Range("A9").End(xlDown).Row
    Sheets("Result").Range("L" & row_result + 1).FormulaR1C1 = "Completed"
    Sheets("Result").Range("L" & row_result + 1).AutoFill Destination:=Sheets("Result").Range("L" & row_result + 1, "L" & row_result2)
     Sheets("Result").Range("A" & row_result + 1, "Y" & row_result2).Interior.Color = RGB(173, 216, 230)
     ' Define the range of the first column
    Set FirstColumn = Sheets("Result").Range("A" & row_result + 1 & ":A" & row_result2) ' Change "A:A" to the column you want to copy from
    
    ' Define the range of the other columns where you want to apply the formatting
    Set OtherColumns = Sheets("Result").Range("J" & row_result + 1 & ":Y" & row_result2) ' Change "R:Z" to the range of columns you want to format
      ' Copy the formatting of the first column and apply it to the other columns
    FirstColumn.copy
    OtherColumns.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear the clipboard
    
    ' Specify the column containing the dates (column A in this example)
    Set dateColumn = Sheets("Result").Range("J" & row_result + 1 & ":J" & row_result2)
    
    ' Change the format of the date column to the desired date format
    dateColumn.NumberFormat = "mm/dd/yyyy hh:mm AM/PM" ' Change the format as needed
    
    ' Optionally, convert the text to actual dates if needed
    dateColumn.Value = dateColumn.Value
    
     If Sheets("Learning completion 2022").Range("N1").Value = "Site" Then
        Sheets("Learning completion 2022").Columns("N:N").Delete Shift:=xlToLeft
    End If
   
End Sub
Sub cap50()

    Sheets("CAP50").Rows("1:1").AutoFilter
    On Error Resume Next
        Sheets("CAP50").AutoFilter.Sort.SortFields.clear
        
    row_initial = Sheets("CAP50").Range("A1").End(xlDown).Row
    
    If Err <> 0 Then
        Sheets("CAP50").Rows("1:1").AutoFilter
    End If
    Err.clear

    Sheets("CAP50").Range("A1", "AB" & row_initial).ClearContents


    Sheets("CAP50_follow_up_source").Rows("20:20").AutoFilter
    On Error Resume Next
        Sheets("CAP50_follow_up_source").AutoFilter.Sort.SortFields.clear
        
    row_com = Sheets("CAP50_follow_up_source").Range("B20").End(xlDown).Row
    
    If Err <> 0 Then
        Sheets("CAP50_follow_up_source").Rows("20:20").AutoFilter
    End If
    Err.clear
    

    
    Sheets("CAP50_follow_up_source").Range("A20", "AC" & row_com).AutoFilter Field:=6, Criteria1:= _
        "Introduction to sustainability and to Valeo's carbon reduction plan (CAP 50)"
    
    Sheets("CAP50_follow_up_source").Range("A20", "AC" & row_com).AutoFilter Field:=10, Criteria1:= _
        "PTS Powertrain Systems"
   
    row_com = Sheets("CAP50_follow_up_source").Range("B20").End(xlDown).Row
        
    Sheets("CAP50_follow_up_source").Range("B20", "AC" & row_com).copy Destination:=Sheets("CAP50").Range("A1")
   
   
   
' data source reset
    Sheets("CAP50_follow_up_source").Rows("20:20").AutoFilter
    


End Sub


Sub type_reset()
    
    row_type = Sheets("Result").Range("C9").End(xlDown).Row
    For i = 9 To row_type
        If Sheets("Result").Range("H" & i).Value = "Event" Then
            Sheets("Result").Range("T" & i).Value = "Event & Session"
            'Sheets("Result").Range("C" & i, "D" & i).Interior.Color = RGB(126, 246, 123)
        ElseIf Sheets("Result").Range("H" & i).Value = "Session" Then
            Sheets("Result").Range("T" & i).Value = "Event & Session"
            'Sheets("Result").Range("C" & i, "D" & i).Interior.Color = RGB(126, 246, 123)
        Else
            Sheets("Result").Range("T" & i).Value = Sheets("Result").Range("H" & i).Value
        End If
    Next
    
End Sub

Sub data_studio_status_reset()
  'after combine and remove the duplicates
  'set pending approval & in progress & not started & completed
    row_result = Sheets("Result").Range("C9").End(xlDown).Row
    For i = 9 To row_result
        If (Sheets("Result").Range("V" & i).Value = "1") Then
            Sheets("Result").Range("U" & i).Value = "Completed"
        ElseIf (Sheets("Result").Range("V" & i).Value = "2") Then
            Sheets("Result").Range("U" & i).Value = "Not Started"
            Sheets("Result").Range("K" & i).Value = "Not Started"
        End If
    Next
End Sub

Public Function Checkpending(str As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.Regexp")
            
    Dim is_exist As Boolean
    With reg
        .Global = True
        .IgnoreCase = True
        .Pattern = "pending"
        is_exist = .test(str)
    End With
    Checkpending = is_exist
End Function

Function set_priority(row_management As Long)
    Dim ss As String
    Dim arr()
    row_result = row_management
    arr = Sheets("Result").Range("A9", "V" & row_result).Value
    
    For i = 1 To row_result - 8
        Application.StatusBar = "Set priority process: " + GetProgress(i, row_result - 8)
        If (arr(i, 12) = "Completed") Then 'L
            Sheets("Result").Range("U" & i + 8).Value = "5.Completed"
            Sheets("Result").Range("V" & i + 8).Value = "1"
        ElseIf (arr(i, 11) = "Registered") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "4.In Progress & Registered" 'Registered
            Sheets("Result").Range("V" & i + 8).Value = "2"
        ElseIf (arr(i, 11) = "Registered / Past Due") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "4.In Progress & Registered"
            Sheets("Result").Range("V" & i + 8).Value = "2"
        ElseIf (arr(i, 11) = "Exception Requested") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "3.Pending Process"
            Sheets("Result").Range("V" & i + 8).Value = "3"
        ElseIf (arr(i, 11) = "Exception Requested / Past Due") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "3.Pending Process"
            Sheets("Result").Range("V" & i + 8).Value = "3"
        ElseIf (arr(i, 11) = "Incomplete") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "2.Incomplete & Event denied"
            Sheets("Result").Range("V" & i + 8).Value = "4"
        ElseIf (arr(i, 11) = "Incomplete / Past Due") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "2.Incomplete & Event denied"
            Sheets("Result").Range("V" & i + 8).Value = "4"
        ElseIf (arr(i, 11) = "In Progress") Then 'K
            Sheets("Result").Range("U" & i + 8).Value = "4.In Progress & Registered"
        ElseIf (arr(i, 11) = "In Progress / Past Due") Then
            Sheets("Result").Range("U" & i + 8).Value = "4.In Progress & Registered"
        ElseIf (arr(i, 11) = "Denied") Then
                If arr(i, 8) = "Event" Then  ' if event denied, not count as needs
                Sheets("Result").Range("U" & i + 8).Value = "2.Incomplete & Event denied"
                Sheets("Result").Range("V" & i + 8).Value = "4"
                Else ' else count as needs
                Sheets("Result").Range("U" & i + 8).Value = "1.Not Started"
                End If
        ElseIf (arr(i, 11) = "Denied / Past Due") Then
                If arr(i, 8) = "Event" Then
                Sheets("Result").Range("U" & i + 8).Value = "2.Incomplete & Event denied"
                Sheets("Result").Range("V" & i + 8).Value = "4"
                Else
                Sheets("Result").Range("U" & i + 8).Value = "1.Not Started"
                End If
        Else
            ss = arr(i, 11)
            Result = Checkpending(ss)
            If (Result = "True") Then
            Sheets("Result").Range("U" & i + 8).Value = "3.Pending Process"
            Sheets("Result").Range("V" & i + 8).Value = "3"
            Else
            Sheets("Result").Range("U" & i + 8).Value = "1.Not Started"
            End If
        End If
    Next
    row_result = Sheets("Result").Range("C9").End(xlDown).Row
    For i = row_management + 1 To row_result
       Sheets("Result").Range("U" & i).Value = "5.Completed"
       Sheets("Result").Range("V" & i).Value = "1"
    Next

End Function

Sub add_training_id()

    Dim arr1()
    row1 = Sheets("Catalog").Range("A2").End(xlDown).Row
    arr1 = Sheets("Catalog").Range("A2", "D" & row1).Value
    
    Dim arr2()
    row2 = Sheets("Result").Range("A8").End(xlDown).Row
    arr2 = Sheets("Result").Range("A9", "Q" & row2).Value
    
    
    For y = 1 To row2 - 8
        Application.StatusBar = "Add_training_id Result process: " + GetProgress(y, row2 - 8)
        For i = 1 To row1 - 1
        If arr1(i, 1) = arr2(y, 7) Then 'title
            If arr1(i, 2) = arr2(y, 17) Then 'provider
                If arr1(i, 3) = arr2(y, 8) Then ' type
                    Sheets("Result").Range("Y" & y + 8).Value = arr1(i, 4)
                ElseIf arr1(i, 3) = "Event" Then
                    If arr2(y, 8) = "Session" Then
                        Sheets("Result").Range("Y" & y + 8).Value = arr1(i, 4)
                    End If
                End If
            End If
        End If
        
        Next
    Next
End Sub

Sub title_reset()

    Dim arr1()
    row1 = Sheets("Catalog").Range("A2").End(xlDown).Row
    arr1 = Sheets("Catalog").Range("A2", "D" & row1).Value
    
    Dim arr2()
    row2 = Sheets("Result").Range("A8").End(xlDown).Row
    arr2 = Sheets("Result").Range("A9", "Y" & row2).Value
    
    
    For y = 1 To row2 - 8
        Application.StatusBar = "Title reset process: " + GetProgress(y, row2 - 8)
        For i = 1 To row1 - 1
        If arr1(i, 4) = arr2(y, 25) Then 'training ID
                    Sheets("Result").Range("G" & y + 8).Value = arr1(i, 1) 'set training title
        End If
        
        Next
    Next
End Sub

'processing
Function GetProgress(curValue, maxValue)
    Dim i As Single, j As Integer, s As String
    i = maxValue / 20
    j = curValue / i
 
    For m = 1 To j
        s = s & " * "
    Next m
    For n = 1 To 20 - j
        s = s & " - "
    Next n
    GetProgress = s & FormatNumber(curValue / maxValue * 100, 2) & "%"
End Function

Function mix11(a As Long, b As Long) As Long
    ' a, b are the line numbers

    Dim ws As Worksheet
    Set ws = Sheets("Result")
    
    ' Initialize variables
    Dim keep As Long
    Dim Delete As Long
    keep = 0
    Delete = 0

    ' Check if a is "Completed"
    If ws.Cells(a, 12).Value = "Completed" Then
        keep = a
        Delete = b
    ElseIf ws.Cells(b, 12).Value = "Completed" Then
        keep = b
        Delete = a
    Else
        ' Check for target time and LM
        If Len(ws.Cells(a, 19).Value) <> 0 Then
            If Len(ws.Cells(b, 19).Value) = 0 Then
                If Len(ws.Cells(a, 11).Value) = 0 Then
                    keep = b
                    Delete = a
                Else
                    Delete = b  ' keep the first lm
                End If
            ElseIf Len(ws.Cells(b, 19).Value) <> 0 Then
                Delete = b  ' keep the first the target time
            End If
        ElseIf Len(ws.Cells(b, 19).Value) <> 0 Then
            keep = a
            Delete = b
        Else
            Delete = b  ' keep the first lm   make sure has sorted by time
        End If
    End If
    
    If keep <> 0 Then
        If (ws.Cells(keep, 12).Value <> "Completed") Then
            'ws.Cells(keep, 18).Value = ws.Cells(Delete, 18).Value
            If (ws.Cells(keep, 19).Value = "") Then
                ws.Cells(keep, 19).Value = ws.Cells(Delete, 19).Value
            End If
        Else
            If (ws.Cells(keep, 19).Value = "") Then
                ws.Cells(keep, 19).Value = ws.Cells(Delete, 19).Value
            End If
        End If
        If Sheets("Result").Cells(keep, 19) = "CVM" Then
           Sheets("Result").Range("A" & keep & ":Y" & keep).Interior.Color = RGB(255, 200, 100)
        End If
    End If
    
    mix11 = Delete
End Function

Sub combine_new()
     'Call MoveLearningCompletionSheet
     'Call UpdateTrainerInformation
     'Call UpdateLearningmanagement
     'Call UpdateCAP
     'Call UpdateAllMyltrainer
     'Call UpdateSessionsfollow
     'Call UpdateCatalog
     
     Application.StatusBar = "Please Wait for the preprocessing to complete"

    'keep the title column and delete the content
    Sheets("Result").Rows("8:8").AutoFilter
    On Error Resume Next
        Sheets("Result").AutoFilter.Sort.SortFields.clear
        
    row_initial = Sheets("Result").Range("C8").End(xlDown).Row
    
    If Err <> 0 Then
        Sheets("Result").Rows("8:8").AutoFilter
    End If
    Err.clear

    
    
    Sheets("Result").Range("A9", "Y" & row_initial).ClearContents
    With Sheets("Result").Range("A9", "Y" & row_initial).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Call trainer_information_copy
    Call learning_management_copy
    Call cvv_copy
    row_management = Sheets("Result").Range("C9").End(xlDown).Row
    Call learning_completion_copy
    Call add_training_id
    Call learning_completions_old
    Call follow_up
    Call cap50
    
    Row_old = Sheets("Result").Range("A8").End(xlDown).Row
    

    'delete duplicate
    Sheets("Result").Range("A8", "Y" & Row_old).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, _
        7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25), Header:=xlYes

     set_priority (row_management) ' set priority to simplify comparison
    
    
     'sort par "Priority" "Assigned Date" "id" "title" "provider"
    Row_total = Sheets("Result").Range("C8").End(xlDown).Row
    
    Sheets("Result").AutoFilter.Sort.SortFields.clear
    ' priority
     Sheets("Result").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("V8", "V" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'assigned date  **the most important**
    Sheets("Result").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("I8", "I" & Row_total), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'people email adress
    Sheets("Result").AutoFilter.Sort.SortFields. _
        clear
    Sheets("Result").AutoFilter.Sort.SortFields. _
        Add Key:=Range("C8", "C" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Training ID
    Sheets("Result").AutoFilter.Sort.SortFields. _
        clear
    Sheets("Result").AutoFilter.Sort.SortFields. _
        Add Key:=Range("Y8", "Y" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    Dim line1 As Long
    Dim line2 As Long
    Dim Index As Long
    
    row_initial = Sheets("Result").Range("C9").End(xlDown).Row
    Application.StatusBar = row_initial
    
    
    line1 = 9 ' Start from row 9
    line2 = line1 + 1
    
    Do While line2 <= row_initial
        Dim Delete As Long
        Delete = 0
    
        ' ID
        If Sheets("Result").Cells(line1, 3).Value = Sheets("Result").Cells(line2, 3).Value Then
            ' Training ID
            If Sheets("Result").Cells(line1, 25).Value = Sheets("Result").Cells(line2, 25).Value And Sheets("Result").Cells(line1, 25).Value <> 0 Then
                Delete = mix11(line1, line2)
            End If
        End If
    
          If Delete <> 0 Then
              'Sheets("Result").Range("A" & Delete, "Y" & Delete).Interior.Color = RGB(255, 0, 0)
               Sheets("Result").Rows(Delete).ClearContents
                If Delete = line1 Then
                    line1 = line2
                    line2 = line2 + 1
                ElseIf Delete = line2 Then
                    line2 = line2 + 1
                End If
                Delete = 0
            Else ' delete = 0 ,  no line to delete
                line1 = line2
                line2 = line2 + 1
            End If
        Loop
              
    
       'sort par "Priority" "Assigned Date" "id" "title" "provider"
    Row_total = Sheets("Result").Range("C8").End(xlDown).Row
    
    Sheets("Result").AutoFilter.Sort.SortFields.clear
    ' priority
     Sheets("Result").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("V8", "V" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'assigned date  **the most important**
    Sheets("Result").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("I8", "I" & Row_total), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'people email adress
    Sheets("Result").AutoFilter.Sort.SortFields. _
        clear
    Sheets("Result").AutoFilter.Sort.SortFields. _
        Add Key:=Range("C8", "C" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Training ID
    Sheets("Result").AutoFilter.Sort.SortFields. _
        clear
    Sheets("Result").AutoFilter.Sort.SortFields. _
        Add Key:=Range("Y8", "Y" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Row_total_new = Sheets("Result").Range("C8").End(xlDown).Row
    
     Call type_reset
    'Call title_reset
     'Call recheck_completion
    Sheets("Num_needs").PivotTables("PivotTable2").ChangePivotCache ActiveWorkbook. _
    PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    Sheets("Result").Range("A8", "Y" & Row_total_new), Version:=6)
  
        
    Sheets("Result").Range("X9", "X" & Row_total_new).FormulaR1C1 = "PTS"
    Sheets("Result").Range("A5").FormulaR1C1 = "=NOW()"
    Sheets("Result").Range("B6").Value = Sheets("Result").Range("A5").Value
    Sheets("Result").Range("Z9").Value = Sheets("Result").Range("A5").Value
 
    Sheets("Result").Range("A5").FormulaR1C1 = ""
    
    Call LM_filter
    Call Session_Completed
    Call CopyRangeToNewSheet
    Call CopySheetToNewWorkbook
    Call CopySheetToNewWorkbookCompleted
    
    Application.StatusBar = False

     
    
End Sub
Sub recheck_completion()

    Dim arr1()
    row_initial = Sheets("Result").Range("C9").End(xlDown).Row
    arr1 = Sheets("Result").Range("A9", "Y" & row_initial).Value
    
    Dim arr2()
    row_com = Sheets("Learning completion 20&21&22").Range("C1").End(xlDown).Row
    arr2 = Sheets("Learning completion 20&21&22").Range("A2", "E" & row_com).Value
    
    For i = 1 To row_initial - 8
        Application.StatusBar = "Recheck completion process: " + GetProgress(i, row_initial - 8)
        For y = 1 To row_com - 1
        
        If arr1(i, 3) = arr2(y, 2) Then  'id
            If arr1(i, 7) = arr2(y, 3) Then 'title
                If arr1(i, 8) = arr2(y, 4) Then 'type
                    If arr1(i, 17) = arr2(y, 5) Then 'provider
                    Sheets("Result").Range("K" & i + 8).Value = "Completed"
                    Sheets("Result").Range("L" & i + 8).Value = "Completed"
                    Sheets("Result").Range("U" & i + 8).Value = "5.Completed"
                    Sheets("Result").Range("V" & i + 8).Value = "1"
                    'Sheets("Result").Range("A" & i + 8, "T" & i + 8).Interior.Color = RGB(105, 182, 74)
                    End If
                Else
                    If arr1(i, 8) = "Event" Then
                        If arr2(y, 4) = "Session" Then
                        Sheets("Result").Range("K" & i + 8).Value = "Completed"
                        Sheets("Result").Range("L" & i + 8).Value = "Completed"
                        Sheets("Result").Range("U" & i + 8).Value = "5.Completed"
                        Sheets("Result").Range("V" & i + 8).Value = "1"
                        End If
                    ElseIf arr1(i, 8) = "Session" Then
                        If arr2(y, 4) = "Event" Then
                        Sheets("Result").Range("K" & i + 8).Value = "Completed"
                        Sheets("Result").Range("L" & i + 8).Value = "Completed"
                        Sheets("Result").Range("U" & i + 8).Value = "5.Completed"
                        Sheets("Result").Range("V" & i + 8).Value = "1"
                        End If
                    End If
                End If
            End If
        End If
        
        
        Next
    Next


End Sub
Sub test_window()

    UserForm1.Show
    
End Sub


Sub CVM_filter_change()

    'clear the old data
    row_initial = Sheets("CVM_filter").Range("A1").End(xlDown).Row
    Sheets("CVM_filter").Range("A2", "S" & row_initial).ClearContents
    With Sheets("CVM_filter").Range("A2", "S" & row_initial).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    'copy the result to cvm_filter
    row_CVM = Sheets("Result").Range("A8").End(xlDown).Row
    Sheets("Result").Range("A8", "T" & row_CVM).AutoFilter Field:=19, Criteria1:="<>"
    
    Sheets("Result").Range("A8", "T" & row_CVM).AutoFilter Field:=12, Criteria1:= _
        "=In Progress", Operator:=xlOr, Criteria2:="=Not Started"
    row_CVM2 = Sheets("Result").Range("A9").End(xlDown).Row
    Sheets("Result").Range("A9", "T" & row_CVM2).copy Destination:=Sheets("CVM_filter").Range("A2")
    
    'reset the filter
    Sheets("Result").Range("A8", "T" & row_CVM).AutoFilter Field:=19
    Sheets("Result").Range("A8", "T" & row_CVM).AutoFilter Field:=12
    
    'refresh the pivot table
    row_cvm_filter = Sheets("CVM_filter").Range("A1").End(xlDown).Row
    Sheets("Master Pivot table (CVM)").PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets("CVM_filter").Range("A1", "T" & row_cvm_filter), Version:=6 _
        )
    
End Sub

Sub add_BG()

    row2 = Sheets("Trainer_information").Range("A1").End(xlDown).Row
    row_bg = Sheets("All_Myl_trainer").Range("A20").End(xlDown).Row
    Sheets("All_Myl_trainer").Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("All_Myl_trainer").Range("C19").FormulaR1C1 = "Full name"
    
    For i = 20 To row_bg
        Sheets("All_Myl_trainer").Range("C" & i).FormulaR1C1 = "=RC[-2]&"", ""&RC[-1]"
    Next
    
    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    For i = 20 To row_bg
        d.Add Sheets("All_Myl_trainer").Range("C" & i).Value, Sheets("All_Myl_trainer").Range("T" & i).Value
    Next
    
    Dim eamil As Object
    Set eamil = CreateObject("scripting.dictionary")
    For i = 20 To row_bg
        eamil.Add Sheets("All_Myl_trainer").Range("C" & i).Value, Sheets("All_Myl_trainer").Range("D" & i).Value
    Next
    
    Sheets("Trainer_information").Range("J1").Value = "BG"
    
    For i = 2 To row2
        'insert the BG
    
        Sheets("Trainer_information").Range("J" & i).Value = d(Sheets("Trainer_information").Range("C" & i).Value)
        If Len(Sheets("Trainer_information").Range("D" & i).Value) = 0 Then
            Sheets("Trainer_information").Range("D" & i).Value = eamil(Sheets("Trainer_information").Range("C" & i).Value)
        End If
    Next

    Sheets("All_Myl_trainer").Columns("C:C").Delete Shift:=xlToLeft

End Sub



Sub follow_up()

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
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=16, Criteria1:=""
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=28, Criteria1:= _
        "Approved"
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=29, Criteria1:=Array( _
        "Approved", "Registered", "Registered / Past Due"), _
        Operator:=xlFilterValues
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).AutoFilter Field:=32, Criteria1:=Array( _
        "Central R&D", "Group R&D", "PowerTECH Knowledge"), Operator:= _
        xlFilterValues
    
    row_initial = Sheets("Follow up list").Range("A1").End(xlDown).Row
    
    Sheets("Follow up list").Range("A1", "AW" & row_initial).ClearContents
    
    Sheets("Sessions follow up source").Range("A1", "AO" & row2).copy Destination:=Sheets("Follow up list").Range("A1")
    
    
    row_cvv = Sheets("Follow up list").Range("A2").End(xlDown).Row
    
    'Sheets("Follow up list").Range("AP1").FormulaR1C1 = "Duration"
    'Sheets("Follow up list").Range("AP2", "AP" & row_cvv).FormulaR1C1 = "=RC[-1]-RC[-2]"
    'Sheets("follow up list").Columns("AP:AP").NumberFormat = "h:mm;@"
    
    
    'Sheets("Follow up list").Range("AQ1").FormulaR1C1 = "Start time CST Mexico City / Detroit Time -8h"
    'Sheets("Follow up list").Range("AQ2", "AQ" & row_cvv).FormulaR1C1 = "=RC[-3]-TIME(8,0,0)"
    
    'Sheets("Follow up list").Range("AR1").FormulaR1C1 = "Start time CEST Paris / Berlin Time"
    'Sheets("Follow up list").Range("AR2", "AR" & row_cvv).FormulaR1C1 = "=RC[-4]"
    
    'Sheets("Follow up list").Range("AS1").FormulaR1C1 = "Start time India Time +3h30m"
    'Sheets("Follow up list").Range("AS2", "AS" & row_cvv).FormulaR1C1 = "=RC[-5]+TIME(3,30,0)"
    
    'Sheets("Follow up list").Range("AT1").FormulaR1C1 = "Start time China Time +6h"
    'Sheets("Follow up list").Range("AT2", "AT" & row_cvv).FormulaR1C1 = "=RC[-6]+TIME(6,0,0)"
    
    'Sheets("Follow up list").Range("AU1").FormulaR1C1 = "Start time Japan/Koreann Time +7h"
    'Sheets("Follow up list").Range("AU2", "AU" & row_cvv).FormulaR1C1 = "=RC[-7]+TIME(7,0,0)"
    
    'Sheets("follow up list").Columns("AQ:AU").NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    Sheets("Sessions follow up source").Rows("1:1").AutoFilter ' filter reset
    
    Call training_parts_details
    
    row_cvv = Sheets("Follow up list").Range("A2").End(xlDown).Row
    Sheets("Follow up list").Columns("O:O").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Follow up list").Columns("O:O").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Follow up list").Columns("O:O").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Follow up list").Columns("O:O").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Sheets("Follow up list").Range("O2", "O" & row_cvv).FormulaR1C1 = "=LEFT(RC[-1],FIND(""@"",RC[-1])-1)"
    Sheets("Follow up list").Range("P2", "P" & row_cvv).FormulaR1C1 = "=LEFT(RC[-1],FIND(""."",RC[-1])-1)"
    Sheets("Follow up list").Range("Q2", "Q" & row_cvv).FormulaR1C1 = "=RIGHT(RC[-2],LEN(RC[-2])-FIND(""."",RC[-2]))"
    Sheets("Follow up list").Range("R2", "R" & row_cvv).FormulaR1C1 = "=IF(ISNUMBER(FIND(""."",RC[-1])),LEFT(RC[-1],FIND(""."",RC[-1])-1),RC[-1])"


    For i = 2 To row_cvv

        Sheets("Follow up list").Range("P" & i).Value = StrConv(Sheets("Follow up list").Range("P" & i).Value, vbProperCase)
        Sheets("Follow up list").Range("R" & i).Value = StrConv(Sheets("Follow up list").Range("R" & i).Value, vbUpperCase)

    Next
    
    Sheets("Follow up list").Columns("Q:Q").Delete Shift:=xlToLeft
    Sheets("Follow up list").Columns("O:O").Delete Shift:=xlToLeft
    Sheets("Follow up list").Range("O1").FormulaR1C1 = "First name"
    Sheets("Follow up list").Range("P1").FormulaR1C1 = "Last name"
    
End Sub

Sub training_parts_details()

    Sheets("Follow up list").Rows("1:1").AutoFilter
    On Error Resume Next
        Sheets("Follow up list").AutoFilter.Sort.SortFields.clear
    If Err <> 0 Then
        Sheets("Follow up list").Rows("1:1").AutoFilter
    End If
    Err.clear
    
    row2 = Sheets("Follow up list").Range("A1").End(xlDown).Row
    
    row_initial = Sheets("Training_parts_details").Range("A1").End(xlDown).Row
    
    Sheets("Training_parts_details").Range("A1", "H" & row_initial).ClearContents
    
    Sheets("Follow up list").Range("H1", "H" & row2).copy Destination:=Sheets("Training_parts_details").Range("A1")
    Sheets("Follow up list").Range("J1", "J" & row2).copy Destination:=Sheets("Training_parts_details").Range("B1")
    Sheets("Follow up list").Range("V1", "V" & row2).copy Destination:=Sheets("Training_parts_details").Range("C1")
    Sheets("Follow up list").Range("W1", "W" & row2).copy Destination:=Sheets("Training_parts_details").Range("D1")
    'Sheets("Follow up list").Range("Q1", "Q" & row2).copy Destination:=Sheets("Training_parts_details").Range("E1")
    'Sheets("Follow up list").Range("AN1", "AN" & row2).copy Destination:=Sheets("Training_parts_details").Range("F1")
    'Sheets("Follow up list").Range("AO1", "AO" & row2).copy Destination:=Sheets("Training_parts_details").Range("G1")
    'Sheets("Follow up list").Range("AP1", "AP" & row2).copy Destination:=Sheets("Training_parts_details").Range("H1")
    'Sheets("Follow up list").Range("AQ1", "AQ" & row2).copy Destination:=Sheets("Training_parts_details").Range("I1")
    'Sheets("Follow up list").Range("AR1", "AR" & row2).copy Destination:=Sheets("Training_parts_details").Range("J1")
    'Sheets("Follow up list").Range("AS1", "AS" & row2).copy Destination:=Sheets("Training_parts_details").Range("K1")
    'Sheets("Follow up list").Range("AT1", "AT" & row2).copy Destination:=Sheets("Training_parts_details").Range("L1")
    'Sheets("Follow up list").Range("AU1", "AU" & row2).copy Destination:=Sheets("Training_parts_details").Range("M1")
    
    Row_total = Sheets("Training_parts_details").Range("A1").End(xlDown).Row
    'start time
    Sheets("Training_parts_details").AutoFilter.Sort.SortFields. _
        clear
    'Sheets("Training_parts_details").AutoFilter.Sort.SortFields. _
        'Add Key:=Range("F1", "F" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending _
        ', DataOption:=xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'local number
    Sheets("Training_parts_details").AutoFilter.Sort.SortFields. _
        clear
    Sheets("Training_parts_details").AutoFilter.Sort.SortFields. _
        Add Key:=Range("D1", "D" & Row_total), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With Sheets("Result").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("Training_parts_details").Range("A1", "M" & Row_total).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), Header:=xlYes
    Sheets("Follow up list").Rows("1:1").AutoFilter ' filter reset

End Sub

Sub LM_filter()

    row_initial = Sheets("LM_filter").Range("A1").End(xlDown).Row
    Sheets("LM_filter").Range("A1", "Y" & row_initial).ClearContents
    With Sheets("LM_filter").Range("A1", "Y" & row_initial).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Sheets("Result").Rows("8:8").AutoFilter
    On Error Resume Next
        Sheets("Result").AutoFilter.Sort.SortFields.clear
        
    If Err <> 0 Then
        Sheets("Result").Rows("8:8").AutoFilter
    End If
    Err.clear
    
    
    Row = Sheets("Result").Range("C8").End(xlDown).Row
    
    Sheets("Result").Range("A8", "Y" & Row).AutoFilter Field:=21, Criteria1:= _
        "1.Not Started"
   
    Sheets("Result").Range("A8", "Y" & Row).copy Destination:=Sheets("LM_filter").Range("A1")
    
    Sheets("Result").Rows("8:8").AutoFilter
    
    
End Sub


