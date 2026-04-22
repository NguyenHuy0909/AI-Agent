Attribute VB_Name = "Mod_func"

Option Explicit

Public Function FindTextCell(ByVal ws As Worksheet, ByVal targetText As String) As Range
    ' Tối ưu: Dùng hàm Find của Excel, siêu tốc độ thay vì quét vòng lặp
    Set FindTextCell = ws.Cells.Find(What:=targetText, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
End Function

Public Function ReadCellToRight(ByVal anchorCell As Range) As String
    If Not anchorCell Is Nothing Then
        ReadCellToRight = Trim$(CStr(anchorCell.Offset(0, 1).value))
    Else
        ReadCellToRight = ""
    End If
End Function

Public Function ReadRowValuesToRight(ByVal anchorCell As Range, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetCol As Long = 1) As String()

    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim startCol As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim values() As String
    Dim cellValue As String
    Dim count As Long

    If anchorCell Is Nothing Then Exit Function
    Set ws = anchorCell.Worksheet
    rowIndex = anchorCell.row + offsetRow
    startCol = anchorCell.Column + offsetCol
    lastCol = ws.Cells(rowIndex, ws.Columns.count).End(xlToLeft).Column

    If lastCol < startCol Then Exit Function
    
    ReDim values(0 To lastCol - startCol)
    count = 0
    For colIndex = startCol To lastCol
        cellValue = Trim$(CStr(ws.Cells(rowIndex, colIndex).value))
        If cellValue <> "" Then
            values(count) = cellValue
            count = count + 1
        End If
    Next colIndex

    If count > 0 Then
        ReDim Preserve values(0 To count - 1)
        ReadRowValuesToRight = values
    End If

End Function

Public Function ReadColumnValuesBelow(ByVal anchorCell As Range, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetCol As Long = 1, Optional ByVal lastRow As Long = 9) As String()

    Dim ws As Worksheet
    Dim dataCol As Long
    Dim startRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim cellValue As String
    Dim count As Long

    If anchorCell Is Nothing Then Exit Function
    Set ws = anchorCell.Worksheet
    dataCol = anchorCell.Column + offsetCol
    startRow = anchorCell.row + offsetRow
    
    ' Sửa lỗi: Khôi phục tham số lastRow (vì có module truyền tham số thứ 4 vào)
    ' Tuy nhiên vẫn tự động dò tìm dòng cuối của khối dữ liệu (dừng lại khi gặp ô trống quá nhiều hoặc dùng lastRow nếu lớn hơn)
    Dim actualLastRow As Long
    actualLastRow = ws.Cells(ws.Rows.count, dataCol).End(xlUp).row
    If lastRow > actualLastRow Then lastRow = actualLastRow
    
    If lastRow < startRow Then Exit Function
    
    ReDim values(0 To lastRow - startRow)
    count = 0
    For rowIndex = startRow To lastRow
        cellValue = Trim$(CStr(ws.Cells(rowIndex, dataCol).value))
        If cellValue <> "" Then
            values(count) = cellValue
            count = count + 1
        End If
    Next rowIndex

    If count > 0 Then
        ReDim Preserve values(0 To count - 1)
        ReadColumnValuesBelow = values
    End If

End Function

Public Function IsStringArrayAllocated(ByRef arr() As String) As Boolean
    On Error Resume Next
    Dim ub As Long
    ub = UBound(arr)
    IsStringArrayAllocated = (Err.Number = 0)
    On Error GoTo 0
End Function
