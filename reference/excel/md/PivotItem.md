# PivotItem

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020876-0000-0000-C000-000000000046}  

Represents an item in a PivotTable field.

**Remarks:** The items are the individual data entries in a field category. The PivotItem object is a member of the PivotItems collection. The PivotItems collection contains all the items in a PivotField object.

**Example:**

```vba
Worksheets("sheet3").PivotTables(1) _
 .PivotFields("year").PivotItems("1998").Visible = False
```

## Properties (22)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As PivotField  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ChildItems As Variant  (read-only)`  
  Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the items (a PivotItems object) that are group children in the specified field, or children of the specified item. Read-only.
- `DataRange As Range  (read-only)`  
  Returns a Range object as shown in the following table. Read-only.
- `_Default As String  (read/write)`
- `LabelRange As Range  (read-only)`  
  Returns a Range object that represents all the cells in the PivotTable report that contain the item. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `ParentItem As PivotItem  (read-only)`  
  Returns a PivotItem object that represents the parent PivotTable item in the parent PivotField object (the field must be grouped so that it has a parent). Read-only.
- `ParentShowDetail As Boolean  (read-only)`  
  True if the specified item is showing because one of its parents is showing detail. False if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped. Read-only Boolean.
- `Position As Long  (read/write)`  
  Returns or sets a Long value that represents the position of the item in its field, if the item is currently showing.
- `ShowDetail As Boolean  (read/write)`  
  True if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/write Variant. For the PivotItem object (or the Range object if the range is in a PivotTable report), this property is set to True if the item is showing detail.
- `SourceName As Variant  (read-only)`  
  Returns a Variant value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.
- `Value As String  (read/write)`  
  Returns or sets a String value that represents the name of the specified item in the PivotTable field.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `IsCalculated As Boolean  (read-only)`  
  True if the PivotTable item is a calculated field or item. Read-only Boolean.
- `RecordCount As Long  (read-only)`  
  Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only Long.
- `Formula As String  (read/write)`  
  Returns or sets a String value that represents the object's formula in A1-style notation and in the language of the macro.
- `Caption As String  (read/write)`  
  Returns a String value that represents the label text for the pivot item.
- `DrilledDown As Boolean  (read/write)`  
  True if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded or visible). Read/write Boolean.
- `StandardFormula As String  (read/write)`  
  Returns or sets a String specifying formulas with standard English (United States) formatting. Read/write.
- `SourceNameStandard As String  (read-only)`  
  Returns a String that represents the PivotTable items' source name in standard English (United States) format settings. Read-only.

## Methods (2)

- `Delete()`  
  Deletes the object.
- `DrillTo(Field As String)`  
  The DrillTo method supports drilling to a specified PivotField from a PivotItem.
