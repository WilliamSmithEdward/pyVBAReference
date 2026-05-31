# CustomView

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024423-0000-0000-C000-000000000046}  

Represents a custom workbook view.

**Remarks:** The CustomView object is a member of the CustomViews collection.

**Example:**

```vba
ThisWorkbook.CustomViews("Current Inventory").Show
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `PrintSettings As Boolean  (read-only)`  
  True if print settings are included in the custom view. Read-only Boolean.
- `RowColSettings As Boolean  (read-only)`  
  True if the custom view includes settings for hidden rows and columns (including filter information). Read-only Boolean.

## Methods (2)

- `Show()`  
  Displays the object.
- `Delete()`  
  Deletes the object.
