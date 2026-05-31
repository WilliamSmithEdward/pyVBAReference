# Styles

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020853-0000-0000-C000-000000000046}  

A collection of all the Style objects in the specified or active workbook.

**Remarks:** Each Style object represents a style description for a range. The Style object contains all style attributes (font, number format, alignment, and so on) as properties. There are several built-in styles-including Normal, Currency, and Percent.

**Example:**

```vba
For i = 1 To ActiveWorkbook.Styles.Count
 Worksheets(1).Cells(i, 1) = ActiveWorkbook.Styles(i).Name
Next
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Style  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Style  (read-only)`

## Methods (2)

- `Add(Name As String, [BasedOn As Variant]) As Style`  
  Creates a new style and adds it to the list of styles that are available for the current workbook.
    - `Name As String` (required): The new style name.
- `Merge(Workbook As Variant) As Variant`  
  Merges the styles from another workbook into the Styles collection.
    - `Workbook As Variant` (required): A Workbook object that represents the workbook that contains styles to be merged.
