Attribute VB_Name = "Mod_ProcessCSV"
Option Explicit

' =======================================================
' Tool: Group CSV files by RPM and SPEC
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
    
    Dim i As Long
    Dim srcFolderPath As String
    Dim srcFolder As Object
    Dim fileObj As Object
    
    Dim fileName As String
    Dim fileExt As String
    
    ' User-defined keywords for classification. Change these values as needed.
    Dim kw1 As String, kw2 As String
    ' === CHANGE KEYWORDS HERE ===
    kw1 = "rpm"   ' Keyword for level 1 folder
    kw2 = "SPEC"  ' Keyword for level 2 folder
    ' ============================
    
    Dim kw1Val As String
    Dim kw2Val As String
    Dim targetDirKw1 As String
    Dim targetDirKw2 As String
    Dim targetFilePath As String
    
    Dim userResponse As VbMsgBoxResult
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
                    ' Inline extraction to remove helper functions
                    kw1Val = "Unknown_" & kw1
                    kw2Val = "Unknown_" & kw2
                    
                    Dim baseName As String
                    Dim parts() As String
                    Dim j As Integer
                    baseName = Left(fileName, InStrRev(fileName, ".") - 1)
                    parts = Split(baseName, "_")
                    
                    For j = 0 To UBound(parts)
                        If InStr(1, LCase(parts(j)), LCase(kw1)) > 0 Then kw1Val = parts(j)
                        If InStr(1, LCase(parts(j)), LCase(kw2)) > 0 Then kw2Val = parts(j)
                    Next j
                    
                    ' Group only if keyword 1 is found
                    If kw1Val <> "" And Not kw1Val Like "Unknown_*" Then
                        ' Create folder gOutputFolderPath\<kw1Val>\
                        targetDirKw1 = fso.BuildPath(Mod_Init.gOutputFolderPath, kw1Val)
                        If Not fso.FolderExists(targetDirKw1) Then
                            fso.CreateFolder targetDirKw1
                        End If
                        
                        ' Create folder gOutputFolderPath\<kw1Val>\<kw2Val>\
                        targetDirKw2 = fso.BuildPath(targetDirKw1, kw2Val)
                        If Not fso.FolderExists(targetDirKw2) Then
                            fso.CreateFolder targetDirKw2
                        End If
                        
                        targetFilePath = fso.BuildPath(targetDirKw2, fileName)
                        
                        ' Handle duplicate files
                        If fso.FileExists(targetFilePath) Then
                            ' If file exists, simply skip it
                        Else
                            ' Normal copy
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
