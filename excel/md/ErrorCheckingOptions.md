# ErrorCheckingOptions

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002445B-0000-0000-C000-000000000046}  

Represents the error-checking options for an application.

**Remarks:** Use the ErrorCheckingOptions property of the Application object to return an ErrorCheckingOptions object. Reference the Item property of the Errors object to view a list of index values associated with error-checking options. After an ErrorCheckingOptions object is returned, you can use the following properties, which are members of the ErrorCheckingOptions object, to set or return error checking options. - BackgroundChecking - EmptyCellReferences - EvaluateToError - InconsistentFormula - IndicatorColorIndex - NumberAsText - OmittedCells - TextDate - UnlockedFormulaCells

**Example:**

```vba
Sub CheckTextDates()

 Dim rngFormula As Range
 Set rngFormula = Application.Range("A1")

 Range("A1").Formula = "'April 23, 00"
 Application.ErrorCheckingOptions.TextDate = True

 ' Perform check to see if 2 digit year TextDate check is on.
 If rngFormula.Errors.Item(xlTextDate).Value = True Then
 MsgBox "The text date error checking feature is enabled."
 Else
 MsgBox "The text date error checking feature is not on."
 End If

End Sub
```

## Properties (15)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BackgroundChecking As Boolean  (read/write)`  
  Alerts the user for all cells that violate enabled error-checking rules. When this property is set to True (default), the AutoCorrect Options button appears next to all cells that violate enabled errors. False disables background checking for errors. Read/write Boolean.
- `IndicatorColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the color of the indicator for error checking options. Read/write XlColorIndex.
- `EvaluateToError As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value. Read/write Boolean.
- `TextDate As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year. Read/write Boolean.
- `NumberAsText As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text. Read/write Boolean.
- `InconsistentFormula As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check. Read/write Boolean.
- `OmittedCells As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells. Read/write Boolean.
- `UnlockedFormulaCells As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas. Read/write Boolean.
- `EmptyCellReferences As Boolean  (read/write)`  
  When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking. Read/write Boolean.
- `ListDataValidation As Boolean  (read/write)`  
  A Boolean value that is True if data validation is enabled in a list. Read/write Boolean.
- `InconsistentTableFormula As Boolean  (read/write)`  
  Returns True if the table formula is inconsistent. Read/write Boolean.
- `MisleadingNumberFormats As Boolean  (read/write)`
