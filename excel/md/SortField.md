# SortField

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A9-0000-0000-C000-000000000046}  

The SortField object contains all the sort information for the Worksheet, ListObject, and AutoFilter objects.

**Remarks:** Developers can use the BeforeSort event to override Excel's default behavior and put their own sort algorithm into the application.

## Properties (11)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `SortOn As XlSortOn  (read/write)`  
  Returns or sets the attribute of the cell to sort on. Read/write XlSortOn.
- `SortOnValue As Object  (read-only)`  
  Returns the value on which the sort is performed for the specified SortField object. Read-only.
- `Key As Range  (read-only)`  
  Specifies the range that is currently being sorted on. Read-only.
- `Order As XlSortOrder  (read/write)`  
  Determines the sort order for the values specified in the key. Read/write.
- `CustomOrder As Variant  (read/write)`  
  Specifies a custom order to sort the fields. Read/write Variant.
- `DataOption As XlSortDataOption  (read/write)`  
  Specifies how to sort text in the range specified in a SortField object. Read/write XlSortDataOption.
- `Priority As Long  (read/write)`  
  Specifies the priority for the sort field. Read/write.
- `SubField As Variant  (read/write)`

## Methods (3)

- `Delete()`  
  Removes the specified SortField object from the SortFields collection.
- `ModifyKey(Key As Range)`  
  Modify the key value by which values are sorted in the field.
    - `Key As Range` (required): Specifies the key to be modified.
- `SetIcon(Icon As Icon)`  
  Sets an icon for a SortField object.
    - `Icon As Icon` (required): The icon to be set.
