# SheetViews

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448C-0000-0000-C000-000000000046}  

A collection of all the sheet views in the specified or active workbook window.

**Example:**

```vba
ActiveWindow.SheetViews.Count
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Item As Object  (read-only)`  
  Returns a SheetView object that represents views in a workbook. Read-only.
- `_Default As Object  (read-only)`

## Methods (1)

- `_NewEnum() As IUnknown`
