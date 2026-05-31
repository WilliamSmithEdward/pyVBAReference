# AutoCorrect

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208D4-0000-0000-C000-000000000046}  

Contains Microsoft Excel AutoCorrect attributes (capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on).

**Example:**

```vba
With Application.AutoCorrect
 .TwoInitialCapitals = True
 .ReplaceText = True
End With
```

## Properties (12)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `CapitalizeNamesOfDays As Boolean  (read/write)`  
  True if the first letter of day names is capitalized automatically. Read/write Boolean.
- `ReplacementList As Variant  (read/write)`  
  Returns the array of AutoCorrect replacements.
- `ReplaceText As Boolean  (read/write)`  
  True if text in the list of AutoCorrect replacements is replaced automatically. Read/write Boolean.
- `TwoInitialCapitals As Boolean  (read/write)`  
  True if words that begin with two capital letters are corrected automatically. Read/write Boolean.
- `CorrectSentenceCap As Boolean  (read/write)`  
  True if Microsoft Excel automatically corrects sentence (first word) capitalization. Read/write Boolean.
- `CorrectCapsLock As Boolean  (read/write)`  
  True if Microsoft Excel automatically corrects accidental use of the CapsLock key. Read/write Boolean.
- `DisplayAutoCorrectOptions As Boolean  (read/write)`  
  Allows the user to display or hide the AutoCorrect Options button. The default value is True. Read/write Boolean.
- `AutoExpandListRange As Boolean  (read/write)`  
  A Boolean value indicating whether automatic expansion is enabled for lists. When you type in a cell of an empty row or column next to a list, the list will expand to include that row or column if automatic expansion is enabled. Read/write Boolean.
- `AutoFillFormulasInLists As Boolean  (read/write)`  
  Affects the creation of calculated columns created by automatic fill-down lists. Read/write Boolean.

## Methods (2)

- `AddReplacement(What As String, Replacement As String) As Variant`  
  Adds an entry to the array of AutoCorrect replacements.
    - `What As String` (required): The text to be replaced. If this string already exists in the array of AutoCorrect replacements, the existing substitute text is replaced by the new text.
    - `Replacement As String` (required): The replacement text.
- `DeleteReplacement(What As String) As Variant`  
  Deletes an entry from the array of AutoCorrect replacements.
    - `What As String` (required): The text to be replaced, as it appears in the row to be deleted from the array of AutoCorrect replacements. If this string doesn't exist in the array of AutoCorrect replacements, this method fails.
