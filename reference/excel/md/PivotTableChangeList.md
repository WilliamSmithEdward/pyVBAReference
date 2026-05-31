# PivotTableChangeList

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C1-0000-0000-C000-000000000046}  

Represents the list of changes a user has made to value cells in a PivotTable report based on an OLAP data source.

**Remarks:** The PivotTableChangeList collection contains ValueChange objects that represent the changes a user has made to value cells in a PivotTable report. To return the PivotTableChangeList collection for a PivotTable, use the ChangeList property of the PivotTable object.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent PivotTable object for the specified PivotTableChangeList collection. Read-only.
- `_Default As ValueChange  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As ValueChange  (read-only)`  
  Returns a single ValueChange object from the specified PivotTableChangeList collection.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (1)

- `Add(Tuple As String, Value As Double, [AllocationValue As Variant], [AllocationMethod As Variant], [AllocationWeightExpression As Variant]) As ValueChange`  
  Adds a ValueChange object to the specified PivotTableChangeList collection.
    - `Tuple As String` (required): The MDX tuple of the value to change in the OLAP data source.
    - `Value As Double` (required): The value to commit.
    - `AllocationValue As Variant` (optional): The value to allocate when performing what-if analysis. If this parameter is not supplied, the default allocation value of the OLAP server will be used.
    - `AllocationMethod As Variant` (optional): The method to use to allocate this value when performing what-if analysis. If this parameter is not supplied, the default allocation method of the OLAP server will be used.
    - `AllocationWeightExpression As Variant` (optional): The MDX weight expression to use for this value when performing what-if analysis. If this parameter is not supplied, the default allocation weight expression of the OLAP server will be used.
