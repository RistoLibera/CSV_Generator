Option Explicit

Function IS_RANGE_VALID(objSheet, rng)
    On Error Resume Next
    IS_RANGE_VALID = (objSheet.Range(rng).Rows.Count >= 1 And objSheet.Range(rng).Columns.Count >= 1)
    If Err.Number <> 0 Then
        IS_RANGE_VALID = False
    End If
End Function

const xlCSV = 6
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim objShell : Set objShell = CreateObject("Wscript.Shell")
Dim objArgs : Set objArgs = WScript.Arguments
Dim argsCount : Set argsCount = objArgs.Count
Dim curDir : curDir = fso.GetAbsolutePathName(".")
Dim strExcelPath
Dim startCell
Dim endCell
Dim sheetName
Dim CSVFileName
Dim re1 : Set re1 = New RegExp
With re1
    .Pattern     = ".*xls.*"
    .IgnoreCase  = True
    .Global      = False
End With
Dim objFolder : Set objFolder = fso.GetFolder(curDir)
Dim allFiles : Set allFiles = objFolder.Files
Dim objFile
Dim dataFile
Dim matchResult

' Get Excel file
For Each objFile in allFiles
    matchResult = re1.Test(objFile.Name)
    If matchResult Then
        dataFile = objFile.Name
        strExcelPath = curDir & "\" & dataFile
    Else
        WScript.Echo "Excel file does not exist!"
        WScript.Quit(1)
    End If
Next

If argsCount <> 4 Then
    WScript.Echo "Should have 4 arguments!"
    WScript.Quit(1)
End If

startCell = objArgs(0)
endCell = objArgs(1)
sheetName = objArgs(2)
CSVFileName = objArgs(3)

' Get Excel data and convert to CSV
Sub outputData()
    Dim objExcel : Set objExcel = CreateObject("Excel.Application")
    Dim objWorkbook : Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)
    Dim objSheet : set objSheet = objExcel.Activeworkbook.Sheets(sheetName)
    Dim dataRange : dataRange = startCell & ":" & endCell
    Dim tempWorkbook

    If Not IS_RANGE_VALID(objSheet, dataRange) Then
        WScript.Echo "Range is invalid!"
        WScript.Quit(1)
    End If

    Set tempWorkbook = objExcel.Workbooks.Add(1)

    objSheet.Range(dataRange).Copy
    tempWorkbook.Worksheets("Sheet1").Range("A1").PasteSpecial
    objExcel.DisplayAlerts = False
    tempWorkbook.SaveAs curDir & "\" & CSVFileName & ".csv", xlCSV
    tempWorkbook.Close
    objExcel.DisplayAlerts = True

    objWorkbook.Close
    objExcel.Application.Quit
    Set objExcel = Nothing

    If Err.Number <> 0 Then
        WScript.Echo Err.Description
        WScript.Quit(1)
    Else
        WScript.Echo "Over!"
        WScript.Quit(0)
    End If
End Sub

Call outputData
