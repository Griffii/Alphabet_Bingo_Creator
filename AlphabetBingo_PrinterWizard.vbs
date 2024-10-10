' Get the full path where the VBS script is located
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the current directory of the VBS file
scriptPath = WScript.ScriptFullName
folderPath = objFSO.GetParentFolderName(scriptPath)

' Construct the full path to the Excel workbook
filePath = folderPath & "\Alphabet_BingoSheet_creator.xlsm"

' Path to the HTA file for selection
htaFilePath = folderPath & "\BingoSelection.hta"

' Path to the temporary file to store the selection from HTA
tempFilePath = folderPath & "\SelectionResult.txt"

' Delete any existing temp file (from a previous run)
If objFSO.FileExists(tempFilePath) Then
    objFSO.DeleteFile(tempFilePath)
End If

' Run the HTA file and wait for it to close
objShell.Run "mshta.exe """ & htaFilePath & """", 1, True

' After HTA closes, read the result from the temporary file
If objFSO.FileExists(tempFilePath) Then
    Set tempFile = objFSO.OpenTextFile(tempFilePath, 1) ' Open for reading
    userChoice = tempFile.ReadLine    ' Read Uppercase or Lowercase
    pageCount = tempFile.ReadLine     ' Read the number of pages to print
    tempFile.Close

    ' Determine which macro to run based on the HTA result
    If LCase(userChoice) = "uppercase" Then
        macroToRun = "CreateAndPrintSixBingoSheetsUppercase"
    ElseIf LCase(userChoice) = "lowercase" Then
        macroToRun = "CreateAndPrintSixBingoSheetsLowercase"
    Else
        MsgBox "No valid selection. Please choose Uppercase or Lowercase."
        WScript.Quit
    End If
Else
    MsgBox "Selection was not made."
    WScript.Quit
End If



' Create an instance of Excel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False ' Set to True if you want to see Excel opening

' Open the workbook
Set objWorkbook = objExcel.Workbooks.Open(filePath)

' Run the selected macro in the workbook multiple times based on the page count
For i = 1 To CInt(pageCount)  ' Convert the pageCount to an integer and loop through
    objExcel.Run macroToRun
Next

' Close the workbook without saving changes
objWorkbook.Close False

' Quit Excel
objExcel.Quit

' Clean up
Set objWorkbook = Nothing
Set objExcel = Nothing