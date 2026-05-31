# Sort

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244AB-0000-0000-C000-000000000046}  

Represents a sort of a range of data.

**Example:**

```vba
Sub SortData()

 'Building data to sort on the active sheet.
 Range("A1").Value = "Name"
 Range("A2").Value = "Bill"
 Range("A3").Value = "Rod"
 Range("A4").Value = "John"
 Range("A5").Value = "Paddy"
 Range("A6").Value = "Kelly"
 Range("A7").Value = "William"
 Range("A8").Value = "Janet"
 Range("A9").Value = "Florence"
 Range("A10").Value = "Albert"
 Range("A11").Value = "Mary"
 MsgBox "The list is out of order. Hit Ok to continue...", vbInformation

 'Selecting a cell within the range.
 Range("A2").Select

 'Applying sort.
 With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
 .SortFields.Clear
 .SortFields.Add Key:=Range("A2:A11"), _
 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
 .SetRange Range("A1:A11")
 .Header = xlYes
 .MatchCase = False
 .Orientation = xlTopToBottom
 .SortMethod = xlPinYin
 .Apply
 End With
 MsgBox "Sort complete.", vbInformation

End Sub
```

## Properties (9)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Rng As Range  (read-only)`  
  Return the range of values on which the sort is performed. Read-only.
- `Header As XlYesNoGuess  (read/write)`  
  Specifies whether the first row contains header information. Read/write XlYesNoGuess.
- `MatchCase As Boolean  (read/write)`  
  Set to True to perform a case-sensitive sort, or set to False to perform a non-case-sensitive sort. Read/write.
- `Orientation As XlSortOrientation  (read/write)`  
  Specifies the orientation for the sort. Read/write XlSortOrientation.
- `SortMethod As XlSortMethod  (read/write)`  
  Specifies the sort method for Chinese languages. Read/write XlSortMethod.
- `SortFields As SortFields  (read-only)`  
  Returns the SortFields object that represents the collection of sort fields associated with the Sort object. Read-only.

## Methods (2)

- `SetRange(Rng As Range)`  
  Sets the range over which the sort occurs.
    - `Rng As Range` (required): Specifies the range over which the sort represented by the Sort object occurs.
- `Apply()`  
  Sorts the range based on the currently applied sort states.
