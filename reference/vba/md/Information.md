# Information

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (14)

- `Err() As ErrObject`  
  Contains information about run-time errors.
- `IMEStatus() As VbIMEStatus`  
  Returns an Integer specifying the current Input Method Editor (IME) mode of Microsoft Windows; available in East Asian versions only.
- `IsArray(VarName As Variant) As Boolean`  
  Returns a Boolean value indicating whether a variable is an array.
- `IsDate(Expression As Variant) As Boolean`  
  Returns True if the expression is a date or is recognizable as a valid date or time; otherwise, it returns False.
- `IsEmpty(Expression As Variant) As Boolean`  
  Returns a Boolean value indicating whether a variable has been initialized.
- `IsError(Expression As Variant) As Boolean`  
  Returns a Boolean value indicating whether an expression is an error value.
- `IsMissing(ArgName As Variant) As Boolean`  
  Returns a Boolean value indicating whether an optional Variant argument has been passed to a procedure.
- `IsNull(Expression As Variant) As Boolean`  
  Returns a Boolean value that indicates whether an expression contains no valid data (Null).
- `IsNumeric(Expression As Variant) As Boolean`  
  Returns a Boolean value indicating whether an expression can be evaluated as a number.
- `IsObject(Expression As Variant) As Boolean`  
  Returns a Boolean value indicating whether an identifier represents an object variable.
- `TypeName(VarName As Variant) As String`  
  Returns a String that provides information about a variable.
- `VarType(VarName As Variant) As VbVarType`  
  Returns an Integer indicating the subtype of a variable, or the type of an object's default property.
- `QBColor(Color As Integer) As Long`  
  Returns a Long representing the RGB color code corresponding to the specified color number.
- `RGB(Red As Integer, Green As Integer, Blue As Integer) As Long`  
  Returns a Long whole number representing an RGB color value.
    - `Red As Integer` (required): Required; Variant (Integer). Number in the range 0&ndash;255, inclusive, that represents the red component of the color.
    - `Green As Integer` (required): Required; Variant (Integer). Number in the range 0&ndash;255, inclusive, that represents the green component of the color.
    - `Blue As Integer` (required): Required; Variant (Integer). Number in the range 0&ndash;255, inclusive, that represents the blue component of the color.
