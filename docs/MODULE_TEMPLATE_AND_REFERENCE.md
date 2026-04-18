# VBA Module Template & Quick Reference
## Use this as a template when creating new VBA modules

---

## VBA MODULE TEMPLATE

```vba
'==============================================================================
' Module Name: ModuleName
' Purpose: [Brief description of what this module does]
' Author: [Your Name]
' Date Created: [Date]
' Version: 1.0
' Last Modified: [Date]
'
' Change Log:
'   1.0 - Initial creation
'   
' Dependencies: [List any modules this depends on]
'
' Usage Example:
'   Dim result As String
'   result = PublicFunctionName("input value")
'   MsgBox result
'==============================================================================

Option Explicit

'===============================================================================
' ENUMS & CONSTANTS
'===============================================================================

' Add module-level constants here
' Example: Const MODULE_VERSION As String = "1.0"

'===============================================================================
' MODULE-LEVEL VARIABLES
'===============================================================================

' Add variables shared across all functions in this module here
' Example: Dim gDebugMode As Boolean


'===============================================================================
' PUBLIC FUNCTIONS - Main Interface
'===============================================================================

'***Function: PublicFunctionName
'Purpose: [What this function does]
'Parameters:
'   param1 (String) - [Description of param1]
'   param2 (Integer) - [Description of param2]
'Returns: Boolean - True if successful, False otherwise
'Example: result = PublicFunctionName("value1", 100)
'***
Public Function PublicFunctionName(param1 As String, param2 As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    ' Input validation
    If param1 = "" Then
        Err.Raise 1001, "PublicFunctionName", "param1 cannot be empty"
    End If
    
    If param2 < 0 Then
        Err.Raise 1002, "PublicFunctionName", "param2 must be non-negative"
    End If
    
    ' Main function logic
    ' TODO: Implement function logic here
    
    PublicFunctionName = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in PublicFunctionName: " & Err.Description, vbCritical, "Module Name"
    PublicFunctionName = False
End Function


'***Function: AnotherPublicFunction
'Purpose: [What this function does]
'Parameters:
'   data (Variant) - [Description]
'Returns: Variant - [Description of return value]
'***
Public Function AnotherPublicFunction(data As Variant) As Variant
    On Error GoTo ErrorHandler
    
    ' Function implementation
    
    Exit Function
ErrorHandler:
    MsgBox "Error in AnotherPublicFunction: " & Err.Description, vbCritical, "Module Name"
End Function


'===============================================================================
' PRIVATE HELPER FUNCTIONS
'===============================================================================

'***Function: HelperFunction
'Purpose: [Internal helper - what it does]
'Parameters:
'   value (String) - [Description]
'Returns: Boolean - [Description]
'Note: Private - only called from within this module
'***
Private Function HelperFunction(value As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Helper logic
    
    Exit Function
ErrorHandler:
    ' Handle error appropriately
    HelperFunction = False
End Function


'===============================================================================
' INITIALIZATION & CLEANUP (if needed)
'===============================================================================

' Optional: Add initialization code here
' Example: When module is first used, set up necessary data structures

' Optional: Add cleanup code here
' Example: When module is no longer needed, free resources

```

---

## QUICK REFERENCE GUIDE

### VBA Syntax Essentials

**Variable Declaration:**
```vba
Dim variableName As DataType
Dim count As Integer
Dim name As String
Dim values() As Double              ' Array
Dim dictionary As Object            ' Generic object
```

**Data Types:**
- `String` - Text
- `Integer` - Whole numbers (-32,768 to 32,767)
- `Long` - Larger whole numbers
- `Double` - Decimal numbers
- `Boolean` - True/False
- `Date` - Date values
- `Object` - Any object reference
- `Variant` - Can hold any type

**Function Declaration:**
```vba
Public Function FunctionName(param1 As String) As String
    ' Code here
    FunctionName = "result"  ' Return value
End Function

Private Sub HelperSub(param1 As Integer)
    ' Sub doesn't return a value
End Sub
```

**Control Structures:**
```vba
' If statement
If condition Then
    ' Code if true
ElseIf otherCondition Then
    ' Code if other condition true
Else
    ' Code if all false
End If

' For loop
For i = 1 To 10
    ' Do something with i
Next i

' For Each loop
For Each item In collection
    ' Do something with item
Next item

' While loop
While condition
    ' Code runs while condition is true
Wend

' Do...Loop
Do While condition
    ' Code runs while condition is true
Loop
```

**Error Handling:**
```vba
On Error GoTo ErrorHandler

' Your code here

Exit Function  ' Skip error handler on success

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Function Name"
    ' Recovery code here
End Sub
```

### Common VBA Operations

**Working with Ranges:**
```vba
' Read value from cell
Dim value As String
value = Range("A1").Value

' Write value to cell
Range("A1").Value = "Hello"

' Read range of cells
Dim data() As Variant
data = Range("A1:C10").Value

' Get active cell
Dim activeCell As Range
Set activeCell = ActiveCell

' Select range
Range("A1:C10").Select
```

**String Operations:**
```vba
Dim str As String
str = "Hello"

' Concatenation
str = str & " World"              ' Result: "Hello World"

' Length
Dim len As Integer
len = Len(str)                    ' Length of string

' Substring
Dim part As String
part = Mid(str, 1, 5)             ' Extract characters

' Case conversion
str = UCase(str)                  ' HELLO WORLD
str = LCase(str)                  ' hello world

' Find text
Dim pos As Integer
pos = InStr(str, "World")         ' Find position of text

' Replace
str = Replace(str, "World", "VBA") ' Replace text
```

**Math Operations:**
```vba
Dim result As Double

result = 5 + 3                    ' Addition: 8
result = 5 - 3                    ' Subtraction: 2
result = 5 * 3                    ' Multiplication: 15
result = 5 / 2                    ' Division: 2.5
result = 5 \ 2                    ' Integer division: 2
result = 5 ^ 2                    ' Exponentiation: 25
result = 5 Mod 3                  ' Modulo (remainder): 2

result = Abs(-5)                  ' Absolute value: 5
result = Int(3.7)                 ' Integer part: 3
result = Round(3.5)               ' Round: 4 (banker's rounding)
result = Sqr(9)                   ' Square root: 3
```

**Date Operations:**
```vba
Dim today As Date
today = Date()                    ' Today's date

Dim now As Date
now = Now()                       ' Current date and time

Dim someDate As Date
someDate = DateValue("01/15/2024")

' Date arithmetic
Dim tomorrow As Date
tomorrow = today + 1              ' Add 1 day

Dim nextMonth As Date
nextMonth = DateAdd("m", 1, today) ' Add 1 month

' Get date parts
Dim year As Integer
year = Year(today)

Dim month As Integer
month = Month(today)

Dim day As Integer
day = Day(today)
```

**Collections & Arrays:**
```vba
' Array declaration
Dim arr(1 To 10) As String        ' Array with 10 elements
Dim arr2() As String              ' Dynamic array

' Add elements
arr(1) = "First"
arr(2) = "Second"

' Array size
Dim size As Integer
size = UBound(arr) - LBound(arr) + 1

' Collection
Dim col As Collection
Set col = New Collection

col.Add "item1"
col.Add "item2"

' Access from collection
Dim item As String
item = col(1)

' Remove from collection
col.Remove 1
```

### Debugging Tips

**Display Debug Information:**
```vba
' Show values in Immediate Window
Debug.Print "Variable value: " & variableName

' Stop execution at breakpoint
Stop  ' Or press F8 to step through code
```

**Common Errors & Solutions:**
- `Type Mismatch` - Assigning wrong type to variable
- `Divide by Zero` - Dividing by 0
- `Object Required` - Using Object when it's Nothing
- `Subscript Out of Range` - Array index out of bounds
- `Runtime Error 1004` - Excel operation failed (often range issue)

---

## MODULE NAMING CONVENTIONS

**Recommended Naming Pattern:**

```
Module Name Structure: [Priority][Category][Purpose]

Examples:
- 00_Config_Constants        (Configuration, highest priority)
- 01_Utility_StringFuncs     (Utilities)
- 02_DataAccess_ExcelRead    (Data access)
- 03_Business_Calculations   (Business logic)
- 04_Integration_UserUI      (Integration/UI)
- 99_Main_Orchestration      (Entry point, lowest priority)
```

**Function Naming Conventions:**

```
Naming Pattern: [Verb][Subject]AsReturnType

Examples:
- ValidateEmailAsBoolean()      (Returns Boolean)
- CalculateTotalAsDouble()      (Returns Double)
- GetActiveSheetAsObject()      (Returns Object)
- FormatDateAsString()          (Returns String)
- ConvertToUpperAsString()      (Returns String)
```

---

## TESTING TEMPLATE FOR AI AGENT

When asking AI Agent to review or generate code:

```
Module Name: [Name]
Purpose: [What module does]

Function List:
1. PublicFunction1(param1 As String) As Boolean
   Purpose: [What it does]
   Error Cases: [What could go wrong]

2. PublicFunction2(data() As String) As Integer
   Purpose: [What it does]
   Error Cases: [What could go wrong]

Questions for AI Agent:
- Does the structure make sense?
- Are there any missing functions?
- Should I use a different approach?
- How should error handling work?

Generate:
- Module template with these functions
- Error handling strategy
- Example usage code
```

---

## FREQUENTLY ASKED VBA QUESTIONS

**Q: How do I reference another module?**
A: Just call the function directly:
```vba
result = OtherModule.PublicFunction(param)
```

**Q: How do I handle errors properly?**
A: Use On Error GoTo with a labeled error handler:
```vba
On Error GoTo ErrorHandler
' Your code
Exit Function
ErrorHandler:
    ' Handle error
End Sub
```

**Q: What's the difference between Public and Private?**
A: Public functions can be called from other modules. Private functions can only be called within the same module.

**Q: How do I work with Excel worksheets?**
A: Use the Sheets collection:
```vba
Sheets("SheetName").Range("A1").Value = "Hello"
Dim ws As Worksheet
Set ws = Sheets("DataSheet")
```

**Q: How do I check if a value exists in an array?**
A: You need to loop through it (VBA lacks built-in contains):
```vba
Function ValueExistsInArray(arr() As String, searchValue As String) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = searchValue Then
            ValueExistsInArray = True
            Exit Function
        End If
    Next i
    ValueExistsInArray = False
End Function
```

---

## NEXT STEPS

1. **Copy this template** when creating new modules
2. **Fill in the blanks** with your module and function information
3. **Share with AI Agent** to get code generation
4. **Follow error handling pattern** from the template
5. **Use naming conventions** from the guide

