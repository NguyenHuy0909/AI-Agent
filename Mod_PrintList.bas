Attribute VB_Name = "Mod_ProcessCSV"
Option Explicit

' =======================================================
' Tool: Group CSV files by keywords (N levels)
' Keywords source : gBodyNames array (from Mod_Init)
' Folder structure: OutputFolder\<kw[0]Val>\<kw[1]Val>\...\<kw[N-1]Val>\
' Skip rule       : skip file if ANY keyword is not found in filename
' Duplicate rule  : skip file if already exists at destination
' =======================================================

Public Sub GroupCSVsByRPM()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure config is loaded
    If Mod_Init.gOutputFolderPath = "" Then
        Call Mod_Init.LoadConfig
    End If

    If Mod_Init.gOutputFolderPath = "" Or Not fso.FolderExists(Mod_Init.gOutputFolderPath) Then
        MsgBox "Output directory is invalid or does not exist: " & Mod_Init.gOutputFolderPath, vbExclamation
        Exit Sub
    End If

    ' Validate gBodyNames array (keyword source)
    If Not Mod_func.IsStringArrayAllocated(Mod_Init.gBodyNames) Then
        MsgBox "gBodyNames array is empty. Please check the #BODY NAME config in the sheet.", vbCritical
        Exit Sub
    End If

    Dim i As Long
    Dim srcFolderPath As String
    Dim srcFolder As Object
    Dim fileObj As Object

    Dim fileName As String
    Dim fileExt As String
    Dim baseName As String

    Dim kwVals() As String
    Dim currentPath As String
    Dim targetFilePath As String
    Dim k As Integer

    Dim fileProcessedCount As Long
    fileProcessedCount = 0

    ' Catch uninitialized array error
    On Error GoTo ErrorHandler
    If UBound(Mod_Init.gSpecFolderPaths) < LBound(Mod_Init.gSpecFolderPaths) Then GoTo ErrorHandler
    On Error GoTo 0

    ' Loop through each Spec folder
    For i = LBound(Mod_Init.gSpecFolderPaths) To UBound(Mod_Init.gSpecFolderPaths)
        srcFolderPath = Mod_Init.gSpecFolderPaths(i)

        If srcFolderPath <> "" And fso.FolderExists(srcFolderPath) Then
            Set srcFolder = fso.GetFolder(srcFolderPath)

            For Each fileObj In srcFolder.Files
                fileName = fileObj.Name
                fileExt = fso.GetExtensionName(fileName)

                If LCase(fileExt) = "csv" Then
                    baseName = Left(fileName, InStrRev(fileName, ".") - 1)

                    ' Extract keyword values from filename.
                    ' Returns unallocated array if ANY keyword is missing (Option B).
                    kwVals = ExtractKeywordValues(baseName)

                    If Mod_func.IsStringArrayAllocated(kwVals) Then
                        ' Build N-level nested folder path from gBodyNames
                        currentPath = Mod_Init.gOutputFolderPath
                        For k = LBound(kwVals) To UBound(kwVals)
                            currentPath = fso.BuildPath(currentPath, kwVals(k))
                            If Not fso.FolderExists(currentPath) Then fso.CreateFolder currentPath
                        Next k

                        ' Copy file — skip if duplicate
                        targetFilePath = fso.BuildPath(currentPath, fileName)
                        If Not fso.FileExists(targetFilePath) Then
                            fso.CopyFile fileObj.Path, targetFilePath, False
                            fileProcessedCount = fileProcessedCount + 1
                        End If
                    End If
                End If
            Next fileObj
        End If
    Next i

    MsgBox "Completed grouping " & fileProcessedCount & " CSV files!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred. gSpecFolderPaths array might be empty.", vbCritical
End Sub

' -------------------------------------------------------
' Extract values matching each keyword in gBodyNames
' from the base filename (parts split by "_").
'
' Returns a filled String array (index = keyword index)
' if ALL keywords are found.
'
' Returns an unallocated array if ANY keyword is missing
' → caller must skip that file (Option B).
' -------------------------------------------------------
Private Function ExtractKeywordValues(ByVal baseName As String) As String()
    Dim keywords() As String
    keywords = Mod_Init.gBodyNames

    Dim results() As String
    ReDim results(LBound(keywords) To UBound(keywords))

    Dim parts() As String
    parts = Split(baseName, "_")

    Dim k As Integer, j As Integer
    For k = LBound(keywords) To UBound(keywords)
        results(k) = ""
        For j = 0 To UBound(parts)
            If InStr(1, LCase(parts(j)), LCase(keywords(k))) > 0 Then
                results(k) = parts(j)
                Exit For
            End If
        Next j
        ' Option B: any keyword not found → return unallocated array (skip file)
        If results(k) = "" Then
            Dim emptyArr() As String
            ExtractKeywordValues = emptyArr
            Exit Function
        End If
    Next k

    ExtractKeywordValues = results
End Function

' =======================================================
' List all files in Output folder (recursive, all levels)
' → Sheet "List" : File Name | Relative Path | Date Modified | Size (KB)
' =======================================================
Public Sub ListAllFilesToSheet()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure config is loaded
    If Mod_Init.gOutputFolderPath = "" Then
        Call Mod_Init.LoadConfig
    End If

    If Mod_Init.gOutputFolderPath = "" Or Not fso.FolderExists(Mod_Init.gOutputFolderPath) Then
        MsgBox "Output directory is invalid or does not exist: " & Mod_Init.gOutputFolderPath, vbExclamation
        Exit Sub
    End If

    ' Get or create "List" sheet
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Sheets("List")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "List"
    End If

    ' Clear existing content
    ws.Cells.Clear
    ws.AutoFilterMode = False

    ' --- Header row ---
    With ws.Cells(1, 1) : .Value = "File Name"       : End With
    With ws.Cells(1, 2) : .Value = "Relative Path"   : End With
    With ws.Cells(1, 3) : .Value = "Date Modified"   : End With
    With ws.Cells(1, 4) : .Value = "Size (KB)"       : End With

    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' --- Scan all files recursively ---
    Dim rowIdx As Long
    rowIdx = 2

    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(Mod_Init.gOutputFolderPath)

    Call ScanFolder(fso, rootFolder, Mod_Init.gOutputFolderPath, ws, rowIdx)

    ' --- Formatting ---
    ws.Columns("A:B").AutoFit
    ws.Columns("C").NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Columns("C").AutoFit
    ws.Columns("D").AutoFit

    ' Zebra stripe rows
    Dim r As Long
    For r = 2 To rowIdx - 1
        If r Mod 2 = 0 Then
            ws.Rows(r).Interior.Color = RGB(235, 241, 251)
        Else
            ws.Rows(r).Interior.Color = RGB(255, 255, 255)
        End If
    Next r

    ' AutoFilter on header
    ws.Range("A1:D1").AutoFilter

    ' Activate sheet
    ws.Activate
    ws.Range("A1").Select

    MsgBox "Listed " & (rowIdx - 2) & " file(s) in sheet 'List'.", vbInformation
End Sub

' -------------------------------------------------------
' Recursive helper: scan all files in a folder and subfolders
' -------------------------------------------------------
Private Sub ScanFolder(ByVal fso As Object, _
                       ByVal folder As Object, _
                       ByVal rootPath As String, _
                       ByVal ws As Worksheet, _
                       ByRef rowIdx As Long)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim relFolder As String

    ' Write each file in current folder
    For Each fileObj In folder.Files
        ' Relative path of the parent folder (excluding output root)
        relFolder = Mid(fileObj.ParentFolder.Path, Len(rootPath) + 2)

        ws.Cells(rowIdx, 1).Value = fileObj.Name
        ws.Cells(rowIdx, 2).Value = relFolder
        ws.Cells(rowIdx, 3).Value = fileObj.DateLastModified
        ws.Cells(rowIdx, 4).Value = Round(fileObj.Size / 1024, 2)

        rowIdx = rowIdx + 1
    Next fileObj

    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        Call ScanFolder(fso, subFolder, rootPath, ws, rowIdx)
    Next subFolder
End Sub
