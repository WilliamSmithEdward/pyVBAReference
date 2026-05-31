# Validation

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002442F-0000-0000-C000-000000000046}  

Represents data validation for a worksheet range.

**Example:**

```vba
Range("e5").Validation _
 .Modify xlValidateList, xlValidAlertStop, "=$A$1:$A$10"
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AlertStyle As Long  (read-only)`  
  Returns the validation alert style. Read-only XlDVAlertStyle.
- `IgnoreBlank As Boolean  (read/write)`  
  True if blank values are permitted by the range data validation. Read/write Boolean.
- `IMEMode As Long  (read/write)`  
  Returns or sets the description of the Japanese input rules. Can be one of the XlIMEMode constants listed in the following table. Read/write Long.
- `InCellDropdown As Boolean  (read/write)`  
  True if data validation displays a drop-down list that contains acceptable values. Read/write Boolean.
- `ErrorMessage As String  (read/write)`  
  Returns or sets the data validation error message. Read/write String.
- `ErrorTitle As String  (read/write)`  
  Returns or sets the title of the data-validation error dialog box. Read/write String.
- `InputMessage As String  (read/write)`  
  Returns or sets the data validation input message. Read/write String.
- `InputTitle As String  (read/write)`  
  Returns or sets the title of the data-validation input dialog box. Read/write String. Limited to 32 characters.
- `Formula1 As String  (read-only)`  
  Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-only String.
- `Formula2 As String  (read-only)`  
  Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween (XlFormatConditionOperator). Can be a constant value, a string value, a cell reference, or a formula. Read-only String.
- `Operator As Long  (read-only)`  
  Returns a Long value that represents the operator for the data validation. See XlFormatConditionOperator.
- `ShowError As Boolean  (read/write)`  
  True if the data validation error message will be displayed whenever the user enters invalid data. Read/write Boolean.
- `ShowInput As Boolean  (read/write)`  
  True if the data validation input message will be displayed whenever the user selects a cell in the data validation range. Read/write Boolean.
- `Type As Long  (read-only)`  
  Returns a Long value containing an XlDVType constant that represents the data type validation for a range.
- `Value As Boolean  (read-only)`  
  Returns a Boolean value that indicates if all the validation criteria are met (that is, if the range contains valid data).

## Methods (3)

- `Add(Type As XlDVType, [AlertStyle As Variant], [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant])`  
  Adds data validation to the specified range.
    - `Type As XlDVType` (required): The validation type.
    - `AlertStyle As Variant` (optional): The validation alert style. Can be one of the following XlDVAlertStyle constants: xlValidAlertInformation, xlValidAlertStop, or xlValidAlertWarning.
    - `Operator As Variant` (optional): The data validation operator. Can be one of the following XlFormatConditionOperator constants: xlBetween, xlEqual, xlGreater, xlGreaterEqual, xlLess, xlLessEqual, xlNotBetween, or xlNotEqual.
    - `Formula1 As Variant` (optional): The first part of the data validation equation. Value must not exceed 255 characters.
    - `Formula2 As Variant` (optional): The second part of the data validation equation when _Operator_ is xlBetween or xlNotBetween (otherwise, this argument is ignored).
- `Delete()`  
  Deletes the object.
- `Modify([Type As Variant], [AlertStyle As Variant], [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant])`  
  Modifies data validation for a range.
    - `Type As Variant` (optional): An XlDVType value that represents the validation type.
    - `AlertStyle As Variant` (optional): An XlDVAlertStyle value that represents the validation alert style.
    - `Operator As Variant` (optional): An XlFormatConditionOperator value that represents the data validation operator.
    - `Formula1 As Variant` (optional): The first part of the data validation equation.
    - `Formula2 As Variant` (optional): The second part of the data validation equation when Operator is xlBetween or xlNotBetween; otherwise, this argument is ignored.
