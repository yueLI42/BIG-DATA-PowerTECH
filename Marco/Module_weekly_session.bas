Attribute VB_Name = "Module1"
Sub Sheet_update()
    Dim ss1 As Object
    Dim pastesheet As Object
    Dim copySheet As Object
    Dim ss2 As Object
    Dim needs As Object
    Dim ui As Object
    Dim response As Variant
    Dim numRow1 As Long
    Dim numRow2 As Long
    Dim start_copy As Variant
    Dim end_copy As Variant
    Dim mark_insert As Long
    Dim add As Long
    Dim mark As Long
    Dim i As Long, j As Long
    Dim source As Object
    Dim destination As Object
    Dim sourceValues As Variant
    Dim leftCell As Object
    Dim currentCell As Object
    Dim bg As String
    Dim lastrow As Long
    Dim range As Object
    Dim conditionalFormatRules As Object
    Dim one_week_after As Date
    Dim next_week As Long
    Dim y As Long
    Dim week_line As Long
    Dim currentDate As Date
    Dim dic As Object
    Dim row1 As Long
    Dim row2 As Long
    Dim iCount As Long
    Dim total As Long
    Dim yes As Long
    Dim no As Long
    Dim flag As Long
    Dim cancel As Long
    Dim result As String
    Dim specs(1 To 4) As Variant
     ' Display a message box with the question
    response = MsgBox("Have you changed the date format?", vbYesNo + vbQuestion, "Date Format Check")
    
    ' Check the user's response
    If response = vbNo Then
        ' If the user clicks "No," do nothing
        Exit Sub
    End If
    ' Set up specs array
    specs(1) = Array("\b(No problem)\b", "#00f017", True)
    specs(2) = Array("\b(With problem)\b", "yellow", True)
    specs(3) = Array("\b(Cancel)\b", "red", True)
    specs(4) = Array("\b(Total num)\b", "blue", True)
    
    ' Get all the values of the sheets we need
    Set ss1 = ThisWorkbook ' Assuming this code is in a VBA module within the same workbook
    Set pastesheet = ss1.Sheets("Sessions organization")
    Set copySheet = ss1.Sheets("Sessions_list")
    Set ss2 = Workbooks.Open("C:\Users\yli6\Downloads\file_to_update\Training Result (Google sheet).xlsx") ' Change the path to your other workbook
    Set needs = ss2.Sheets("Num_needs")

    ' Get the last rows
    numRow1 = pastesheet.Cells(pastesheet.Rows.Count, 3).End(xlUp).row
    numRow2 = copySheet.Cells(copySheet.Rows.Count, 2).End(xlUp).row

    pastesheet.range("H3").Value = "Updating...Please wait..."
    pastesheet.range("H3").Interior.Color = RGB(255, 0, 0)

    mark_insert = 0
    add = 0
    mark = 0
    flag = 0

    ' Split the copy sheet
    If copySheet.Cells(1, 12).Value = "" Then
        copySheet.Cells(1, 3).EntireColumn.Insert
        copySheet.Cells(1, 5).EntireColumn.Insert
    End If
    If copySheet.Cells(1, 6).Value <> "" Then
        copySheet.Cells(1, 6).EntireColumn.Insert
    End If

    ' Split the date for start date and end date
    For i = 2 To numRow2
        start_copy = copySheet.Cells(i, 2).Value
        end_copy = copySheet.Cells(i, 4).Value

        Dim start_copy_parts As Variant
        Dim end_copy_parts As Variant

        start_copy_parts = Split(start_copy, vbLf)
        end_copy_parts = Split(end_copy, vbLf)

        Dim valuesToUpdate As Variant
        ReDim valuesToUpdate(1 To 1, 1 To 4)

        valuesToUpdate(1, 1) = start_copy_parts(0)
        valuesToUpdate(1, 2) = start_copy_parts(1)
        valuesToUpdate(1, 3) = end_copy_parts(0)
        valuesToUpdate(1, 4) = end_copy_parts(1)

        copySheet.Cells(i, 2).Resize(1, 4).Value = valuesToUpdate
    Next i

    ' Move the title column to the beginning
    copySheet.Columns(8).Cut
    copySheet.Columns(2).Insert Shift:=xlToRight

    ' Update the range for loop, reduce the number of looping
    ' Loop from Yesterday to the end of the year
    For i = 8 To numRow1
        If pastesheet.Cells(i, 3).Value <= GetYesterday() Then
            numRow1 = i
            Exit For
        End If
        If pastesheet.Cells(i, 4).Value <> "" Then
            pastesheet.Cells(i, 2).Interior.Color = RGB(255, 0, 0)
        End If
    Next i

    For i = 8 To numRow1
        ' Clear the date color that has changed
        If pastesheet.Cells(i, 3).Interior.Color = RGB(167, 86, 222) Then
            pastesheet.Cells(i, 3).Interior.Color = pastesheet.Cells(i, 7).Interior.Color
        End If
    Next i

    ' Compare and update the sessions
    For i = 2 To numRow2
        mark = 0
        For j = 8 + add To numRow1 + add
            If mark_insert = 1 Then
                add = add + 1 ' Number of new rows
                mark_insert = 0
            End If
            ' Compare the locator number, the same session, no need to insert a new row
            If copySheet.Cells(i, 9).Value = pastesheet.Cells(j, 9).Value Then
                mark = 1
                ' Copy and set new values
                ' Only copy the value without the formula
                ' If the date changes, mark the date
                ' Need to transfer in text
                If pastesheet.Cells(j, 3).Value <> copySheet.Cells(i, 3).Value Then
                    flag = 1
                End If
                pastesheet.Cells(j, 2).Interior.Color = pastesheet.Cells(j, 7).Interior.Color
                copySheet.Cells(i, 1).Resize(1, 6).Copy pastesheet.Cells(j, 1)
                copySheet.Cells(i, 8).Resize(1, 5).Copy pastesheet.Cells(j, 8)
                If flag = 1 Then
                    pastesheet.Cells(j, 3).Interior.Color = RGB(167, 86, 222)
                    flag = 0
                End If
                Exit For
            End If
        Next j
        ' New session, insert a new row
        If mark = 0 Then
            Set source = copySheet.Cells(i, 1).Resize(1, 12) ' Copy row i, columns 1-12
            pastesheet.Rows(8).Insert Shift:=xlDown ' Insert one row at line 8
            pastesheet.Cells(8, 1).Resize(1, 20).Interior.Color = RGB(255, 255, 255)
            Set destination = pastesheet.Cells(8, 1).Resize(1, 12)
            sourceValues = source.Value ' Get the values from the source range
            destination.Value = sourceValues ' Paste the values only
            ' Get the left cell's format
            Set leftCell = pastesheet.Cells(8, 1)
            ' Set the current cell's format to match the left cell
            Set currentCell = pastesheet.Cells(8, 2) ' The cell in column 2
            leftCell.Copy
            currentCell.PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False ' Clear clipboard
            ' Set color for new session in the name cell
            pastesheet.Cells(8, 2).Interior.Color = RGB(255, 0, 255)
            mark_insert = 1
        End If
    Next i

    ' Sort by date descending
    pastesheet.range("C7").Sort Key1:=pastesheet.range("C8"), Order1:=xlDescending, Header:=xlYes

    Insert_week pastesheet

    ' Update the range for loop, reduce the number of looping again
    row1 = pastesheet.Cells(pastesheet.Rows.Count, 3).End(xlUp).row
    For i = 8 To row1
        If pastesheet.Cells(i, 3).Value <= GetYesterday() Then
            row1 = i ' The line of yesterday
            Exit For
        End If
    Next i

    Insert_num_training_needs needs, pastesheet, row1
    
    ' Close the external workbook without saving changes
        ss2.Close SaveChanges:=False
           
    Call add_link
    
    ' Insert the updating date
    Dim now As Date
    now = Date
    pastesheet.range("H2").Value = "Last updating date:"
    pastesheet.range("H3").Value = now
    pastesheet.range("H3").Interior.Color = RGB(81, 234, 81)

    MsgBox "Update successfully!"
End Sub



Function GetYesterday() As Date
    Dim today As Date
    Dim yesterday As Date
    today = Date
    yesterday = DateAdd("d", -1, today)
    GetYesterday = yesterday
End Function

Function Get_one_week_after(i As Long, source As Object) As Date
    Dim dateStr As String
    Dim today As Date
    Dim day_after As Date
    Dim MILLIS_PER_DAY As Long
    dateStr = source.Cells(i, 3).Value
    today = CDate(dateStr)
    MILLIS_PER_DAY = 1000 * 60 * 60 * 24
    day_after = DateAdd("d", 7, today)
    Get_one_week_after = day_after
End Function

Sub Insert_week(source As Object)
    Dim one_week_after As Date
    Dim next_week As Long
    Dim y As Long
    Dim week_line As Long
    Dim currentDate As Date
    Dim i As Long
    Dim lastline As Long
    lastline = source.Cells(source.Rows.Count, 1).End(xlUp).row
    For i = 8 To lastline
        If Judge_week(i, source) Then
            Exit For
        End If
    Next i
    While y > 7
        one_week_after = Get_one_week_after(week_line, source)
        next_week = CInt(Mid(source.Cells(week_line, 1).Value, 6)) + 1
        If source.Cells(y, 3).Value > one_week_after Then
            source.Rows(y + 1).Insert Shift:=xlDown
            source.Cells(y + 1, 1).Value = "WEEK " & next_week
            source.Cells(y + 1, 3).Value = one_week_after
            source.Cells(y + 1, 1).Resize(1, 18).Interior.Color = RGB(0, 176, 240)
            week_line = y + 1
        Else
            y = y - 1
        End If
    Wend
    If Not Judge_week(8, source) Then
        one_week_after = Get_one_week_after(week_line, source)
        source.Rows(8).Insert Shift:=xlDown
        source.Cells(8, 1).Value = "WEEK " & next_week
        source.Cells(8, 3).Value = one_week_after
        source.Cells(8, 1).Resize(1, 18).Interior.Color = RGB(0, 176, 240)
        currentDate = Get_one_week_after(8, source)
    End If
End Sub

Sub Insert_num_training_needs(needs As Object, pastesheet As Object, row2 As Long)
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 5 To needs.Cells(needs.Rows.Count, 1).End(xlUp).row
        dic(LCase(Replace(Trim(needs.Cells(i, 1).Value), " ", ""))) = needs.Cells(i, 2).Value
    Next i
    Dim y As Long
    For y = 8 To row2
        pastesheet.Cells(y, 13).Value = dic(LCase(Replace(Trim(pastesheet.Cells(y, 2).Value), " ", "")))
    Next y
End Sub



Sub FormatText(ByRef range As Object, specs() As Variant)
    Dim values As Variant
    Dim match As Object
    Dim formattedText As Variant
    Dim row As Long
    Dim col As Long
    Dim spec As Variant
    ReDim formattedText(1 To range.Rows.Count, 1 To range.Columns.Count)

    values = range.Value
    For row = 1 To UBound(values, 1)
        For col = 1 To UBound(values, 2)
            formattedText(row, col) = values(row, col)
            For Each spec In specs
                Set match = CreateObject("VBScript.RegExp")
                match.Global = True
                match.IgnoreCase = True
                match.Pattern = spec(0)
                If match.Test(values(row, col)) Then
                    formattedText(row, col) = Replace(formattedText(row, col), match, "")
                    formattedText(row, col).Characters(match.FirstIndex + 1, match.Length).Font.Color = RGB(0, 0, 0)
                    formattedText(row, col).Characters(match.FirstIndex + 1, match.Length).Font.Bold = spec(2)
                    formattedText(row, col).Characters(match.FirstIndex + 1, match.Length).Font.Color = RGB(0, 0, 0)
                End If
            Next spec
        Next col
    Next row
    range.Value = formattedText
End Sub

Function Judge_week(i As Long, source As Object) As Boolean
    Dim week As String
    Dim regExp As Object
    Dim res As Variant
    week = source.Cells(i, 1).Value
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = "^WEEK"
    Set res = regExp.Execute(week)
    If Not res Is Nothing Then
        Judge_week = True
    Else
        Judge_week = False
    End If
End Function
Sub add_link()
    Set ss1 = ThisWorkbook ' Assuming this code is in a VBA module within the same workbook
    Set pastesheet = ss1.Sheets("Sessions organization")
    Set copySheet = ss1.Sheets("Sessions_list")
    Set ss2 = Workbooks.Open("C:\Users\yli6\Downloads\source\[CATALOG]_!_Full_MyLearning_Catalog_WITH_PREFERRED_TRAINERS_!.xlsx") ' Change the path to your other workbook
    Set linksheet = ss2.Sheets(1)
    Dim searchString As String
    Dim end_line As Long
     ' Update the range for loop, reduce the number of looping
    end_line = pastesheet.Cells(pastesheet.Rows.Count, 3).End(xlUp).row
    For i = 8 To end_line
        If pastesheet.Cells(i, 3).Value <= GetYesterday() Then
            end_line = i ' The line of yesterday
            Exit For
        End If
    Next i
    Dim lastrow As Long
    lastrow = linksheet.Cells(linksheet.Rows.Count, "A").End(xlUp).row
    
    For i = 9 To end_line
        If pastesheet.Cells(i, 20) = 0 Then
               searchString = pastesheet.Cells(i, 2)
             ' Loop through each cell in Column A
                For j = 1 To lastrow
                    ' Check if the cell in Column A contains the target string (case-insensitive)
                    If InStr(1, linksheet.Cells(j, 1).Value, searchString, vbTextCompare) > 0 Then
                      pastesheet.Cells(i, 20) = linksheet.Cells(j, 27)
                      pastesheet.Cells(i, 20).Font.Color = RGB(0, 0, 255) ' Blue font color (you can change the color)
                      pastesheet.Cells(i, 20).Font.Underline = xlUnderlineStyleSingle ' Underline the text
                    End If
                Next j
            
        End If
    Next i
  
     ' Close the external workbook without saving changes
        ss2.Close SaveChanges:=False
End Sub

