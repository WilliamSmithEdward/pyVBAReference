# CubeFields

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002444D-0000-0000-C000-000000000046}  

A collection of all CubeField objects in a PivotTable report that is based on an OLAP cube. Each CubeField object represents a hierarchy or measure field from the cube.

**Example:**

```vba
Set objNewSheet = Worksheets.Add
intRow = 1
For Each objCubeFld In _
 Worksheets("Sheet1").PivotTables(1).CubeFields
 If objCubeFld.Orientation = xlDataField Then
 objNewSheet.Cells(intRow, 1).Value = objCubeFld.Name
 intRow = intRow + 1
 End If
Next objCubeFld
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
- `Item As CubeField  (read-only)`  
  Returns a single object from a collection.
- `_Default As CubeField  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `AddSet(Name As String, Caption As String) As CubeField`  
  Adds a new CubeField object to the CubeFields collection. The CubeField object corresponds to a set defined on the Online Analytical Processing (OLAP) provider for the cube.
    - `Name As String` (required): A valid name in the SETS schema rowset.
    - `Caption As String` (required): A string representing the field that will be displayed in the PivotTable view.
- `GetMeasure(AttributeHierarchy As Variant, Function As XlConsolidationFunction, [Caption As Variant]) As CubeField`  
  Given an attribute hierarchy, returns an implicit measure for the given function that corresponds to this attribute. If an implicit measure does not exist, a new implicit measure is created and added to the CubeFields collection.
    - `AttributeHierarchy As Variant` (required): The unique cube field that is an attribute hierarchy (XlCubeFieldType = xlHierarchy, and XlCubeFieldSubType = xlCubeAttribute).
    - `Function As XlConsolidationFunction` (required): The function performed in the added data field.
    - `Caption As Variant` (optional): The label used in the PivotTable report to identify this measure. If the measure already exists, _Caption_ will overwrite the existing label of this measure.
