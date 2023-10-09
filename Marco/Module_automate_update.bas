Attribute VB_Name = "Module1"
Sub RenameFiles()
    Dim SourceFolder As String
    Dim File As Object
    Dim NewFileName As String
    Dim RegEx As Object
    
    ' Create a regular expression object
    Set RegEx = CreateObject("VBScript.RegExp")
    
    ' Define the regular expression pattern to match the timestamp
    RegEx.Pattern = "_\d{8}_\d{2}_\d{2}_\d{2}_\w{2}\.xlsx"
    
    ' Define the source folder path
    SourceFolder = "C:\Users\yli6\Downloads\source\"

    ' Loop through each file in the folder
    For Each File In CreateObject("Scripting.FileSystemObject").GetFolder(SourceFolder).Files
        If Not RegEx.Test(File.Name) Then ' Check if the file name does not match the pattern
            ' Delete the file
            Kill SourceFolder & File.Name
        Else
            ' Generate the new file name by removing the timestamp
            NewFileName = RegEx.Replace(File.Name, ".xlsx")
            
            ' Rename the file
            Name SourceFolder & File.Name As SourceFolder & NewFileName
        End If
    Next File

    ' Display a message when the renaming is complete
    'MsgBox "File renaming completed."
End Sub

Sub update()

   Call RenameFiles
   Call OpenFileAndExecuteMacro
   MsgBox "File updating completed."
End Sub

Sub OpenFileAndExecuteMacro()
    Dim wb As Workbook
    Dim filePath As String
    Dim macroName As String
    
    ' Set the file path of the workbook you want to open
    filePath = "C:\Users\yli6\Downloads\file_to_update\Traning Result new version2.0.xlsm" ' Change this to the actual file path
    
    ' Set the name of the macro you want to execute
    macroName = "combine_new" ' Change this to the actual macro name
    
    ' Error handling in case the file doesn't exist
    On Error Resume Next
    
    ' Open the workbook
    Set wb = Workbooks.Open(filePath)
    
    ' Check if the workbook was opened successfully
    If wb Is Nothing Then
        MsgBox "File not found or could not be opened.", vbExclamation
        Exit Sub
    End If
    
    ' Call the macro in the opened workbook
    Application.Run "'" & wb.Name & "'!" & macroName
    
    ' Close the opened workbook (if needed)
    wb.Close SaveChanges:=True
    
    ' Clean up
    Set wb = Nothing
    
    ' Reset error handling
    On Error GoTo 0
End Sub
