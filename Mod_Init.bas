Attribute VB_Name = "Mod_Init"

Option Explicit

' ===== Global config =====
Public gTemplatePath As String
Public gSpecFolderPaths() As String
Public gBodyNames() As String
Public gResultNames() As String
Public gResultMarkers() As String
Public gOutputSheetNames() As String
Public gNamingRuleValues() As String
Public gOutputFolderPath As String

Public Sub LoadConfig()

    Dim wsConfig As Worksheet
    Dim markerCell As Range

    Set wsConfig = ThisWorkbook.Sheets(1)

    Set markerCell = FindTextCell(wsConfig, "#TEMPLATE FILE PATH")
    gTemplatePath = ReadCellToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#SPEC. FOLDER")
    gSpecFolderPaths = ReadColumnValuesBelow(markerCell, 0, 1, 11)

    Set markerCell = FindTextCell(wsConfig, "#BODY NAME")
    gBodyNames = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#RESULT NAME")
    gResultNames = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#RESULT MARKER")
    gResultMarkers = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#SHEET GROUPS")
    gOutputSheetNames = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#NAMING RULE")
    gNamingRuleValues = ReadRowValuesToRight(markerCell, 1)

    Set markerCell = FindTextCell(wsConfig, "#OUTPUT DIRECTORY")
    gOutputFolderPath = ReadCellToRight(markerCell)

End Sub
