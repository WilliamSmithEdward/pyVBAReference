# PivotLineCells

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447F-0000-0000-C000-000000000046}  

Collection of PivotCell objects for a specific PivotLine.

**Remarks:** Use the PivotLineCells (_index_) property of the PivotLine object to return or specify the location of a specific PivotCell object in the collection. You can also specify a PivotField object or the PivotField name to return a single PivotCell object.

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified PivotLineCells object. Read-only.
- `_Default As PivotCell  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As PivotCell  (read-only)`  
  Returns a specific element of the PivotLineCells collection object by its position in the collection. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of items in the PivotLineCells collection. Read-only.
- `Full As Boolean  (read-only)`  
  When True, all of the PivotLineCells in the PivotLine (including those that are hidden in compact form) are retrieved. Read-only Boolean.
