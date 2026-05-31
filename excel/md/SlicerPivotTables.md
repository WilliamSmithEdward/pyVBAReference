# SlicerPivotTables

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244CB-0000-0000-C000-000000000046}  

Represents information about the collection of PivotTables associated with the specified SlicerCache object.

**Remarks:** The SlicerPivotTables collection contains information about the PivotTables that the slicer cache is currently filtering. It provides properties for determining the number of PivotTables that the slicer is associated with, and for retrieving PivotTable objects that represent the PivotTables being filtered. It also provides methods for adding and removing PivotTables from the SlicerPivotTables collection. The SlicerPivotTables collection will be empty if the slicer associated with the specified SlicerCache is not connected to any PivotTables. Use the PivotTables property of the SlicerCache object to return the SlicerPivotTables collection associated with a SlicerCache, which in turn may be associated with one or more slicers.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the SlicerCache object that is the parent of the specified SlicerPivotTables collection. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection.
- `Item As PivotTable  (read-only)`  
  Returns a single PivotTable object from the collection.
- `_Default As PivotTable  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `AddPivotTable(PivotTable As PivotTable)`  
  Adds a reference to a PivotTable to the SlicerPivotTables collection.
    - `PivotTable As PivotTable` (required): A PivotTable object that represents the PivotTable to add.
- `RemovePivotTable(PivotTable As Variant)`  
  Removes a reference to a PivotTable from the SlicerPivotTables collection.
    - `PivotTable As Variant` (required): A PivotTable object that represents the PivotTable to remove, or the name or index of the PivotTable in the collection.
