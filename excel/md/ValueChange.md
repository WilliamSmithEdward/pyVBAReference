# ValueChange

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C0-0000-0000-C000-000000000046}  

Represents a value that has been changed in a PivotTable report that is based on an OLAP data source.

**Remarks:** The PivotTableChangeList collection contains ValueChange objects that represent changes that a user has made to value cells in a PivotTable report. The properties of the ValueChange object specify details about the change that was made, such as the value of the change, the tuple associated with the cell that was changed, the order in which the change was made relative to other changes, and whether the cell is visible in the PivotTable. The ValueChange object also provides the PivotCell property that returns a PivotCell object that represents the cell that was changed, and provides additional information about the changed cell.

## Properties (11)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Order As Long  (read-only)`  
  Returns a value that indicates the order in which this change was performed relative to other changes in the PivotTableChangeList collection. Read-only.
- `VisibleInPivotTable As Boolean  (read-only)`  
  Returns whether the cell (tuple) is currently visible in the PivotTable report. Read-only.
- `PivotCell As PivotCell  (read-only)`  
  Returns a PivotCell object that represents the cell (tuple) that was changed. Read-only.
- `Tuple As String  (read-only)`  
  Returns the MDX tuple of the value that was changed in the OLAP data source. Read-only.
- `Value As Double  (read-only)`  
  Returns the value that the user entered in the cell or that the formula in the cell was evaluated to when the UPDATE CUBE statement was last run against the OLAP data source. Read-only.
- `AllocationValue As XlAllocationValue  (read-only)`  
  Returns the value to allocate when performing what-if analysis. Read-only.
- `AllocationMethod As XlAllocationMethod  (read-only)`  
  Returns the method to use to allocate this value when performing what-if analysis. Read-only.
- `AllocationWeightExpression As String  (read-only)`  
  Returns the MDX weight expression to use for this value when performing what-if analysis. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified ValueChange object from the PivotTableChangeList collection.
