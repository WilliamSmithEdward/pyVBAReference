# Slicers

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C7-0000-0000-C000-000000000046}  

A collection of Slicer objects.

**Remarks:** Each Slicer object represents a slicer in a workbook. Slicers are used to filter data.

**Example:**

```vba
MsgBox ActiveWorkbook.SlicerCaches(1).Slicers.Count
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent SlicerCache object for the specified Slicers collection. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection.
- `Item As Slicer  (read-only)`  
  Returns a Slicer object from the collection. Read-only.
- `_Default As Slicer  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(SlicerDestination As Variant, [Level As Variant], [Name As Variant], [Caption As Variant], [Top As Variant], [Left As Variant], [Width As Variant], [Height As Variant]) As Slicer`  
  Creates a new slicer and returns a Slicer object.
    - `SlicerDestination As Variant` (required): A String that specifies the name of the sheet, or a Worksheet object that represents the sheet, where the resulting slicer will be placed. The destination sheet must be in the workbook that contains the Slicers object specified by expression.
    - `Level As Variant` (optional): For OLAP data sources, the ordinal or the Multidimensional Expression (MDX) name of the level on which the slicer creation is based. Not supported for non-OLAP data sources.
    - `Name As Variant` (optional): The name of the slicer. Excel automatically generates a name if one is not specified. The name must be unique across all slicers within a workbook.
    - `Caption As Variant` (optional): The caption of the slicer.
    - `Top As Variant` (optional): The initial vertical position of the slicer, in points, relative to the upper-left corner of cell A1 on a worksheet.
    - `Left As Variant` (optional): The initial horizontal position of the slicer, in points, relative to the upper-left corner of cell A1 on a worksheet.
    - `Width As Variant` (optional): The initial width, in points, of the slicer control.
    - `Height As Variant` (optional): The initial height, in points, of the slicer control.
