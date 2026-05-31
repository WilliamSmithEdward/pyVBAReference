# _HiddenModule

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (6)

- `Array(ArgList As SAFEARRAY(Variant)) As Variant`  
  Returns a Variant containing an array.
- `_B_str_InputB(Number As Long, FileNumber As Integer) As String`
- `_B_var_InputB(Number As Long, FileNumber As Integer) As Variant`
- `_B_str_Input(Number As Long, FileNumber As Integer) As String`
- `_B_var_Input(Number As Long, FileNumber As Integer) As Variant`
- `Width(FileNumber As Integer, Width As Integer)`  
  Assigns an output line width to a file opened by using the Open statement.
    - `FileNumber As Integer` (required): Required. Any valid file number.
    - `Width As Integer` (required): Required. Numeric expression in the range 0&ndash;255, inclusive, that indicates how many characters appear on a line before a new line is started. If _width_ equals 0, there is no limit to the length of a line. The default value for _width_ is 0.
