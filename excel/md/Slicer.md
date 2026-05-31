# Slicer

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244C8-0000-0000-C000-000000000046}  

Represents a slicer in a workbook.

**Remarks:** Each Slicer object represents a slicer in a workbook. Slicers are used to filter data in PivotTable reports or OLAP data sources. Use the Add method to add a Slicer object to the Slicers collection. To access the SlicerItem object that represents the currently selected button in a slicer, use the ActiveItem property of the Slicer object.

**Example:**

```vba
ActiveWorkbook.SlicerCaches(1).Slicers(1).Caption = "My Slicer"
```

## Properties (22)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the Worksheet object that represents the sheet that contains the slicer. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified slicer. Read/write.
- `Caption As String  (read/write)`  
  Returns or sets the caption of the specified slicer. Read/write.
- `Top As Double  (read/write)`  
  Returns or sets the vertical position of the specified slicer, in points, relative to the upper-left corner of cell A1 on a worksheet. Read/write.
- `Left As Double  (read/write)`  
  Returns or sets the horizontal position of the specified slicer, in points, relative to the upper-left corner of cell A1 on a worksheet. Read/write.
- `DisableMoveResizeUI As Boolean  (read/write)`  
  Returns or sets whether the specified slicer can be moved or resized by using the user interface. Read/write.
- `Width As Double  (read/write)`  
  Returns or sets the width of the specified slicer, in points. Read/write.
- `Height As Double  (read/write)`  
  Returns or sets the height of the specified slicer, in points. Read/write.
- `RowHeight As Double  (read/write)`  
  Returns or sets the height, in points, of each row in the specified slicer. Read/write.
- `ColumnWidth As Double  (read/write)`  
  Returns or sets the width, in points, of each column in the slicer. Read/write.
- `NumberOfColumns As Long  (read/write)`  
  Returns or sets the number of columns in the specified slicer. Read/write.
- `DisplayHeader As Boolean  (read/write)`  
  Returns or sets whether the header that displays the slicer Caption property is visible. Read/write.
- `Locked As Boolean  (read/write)`  
  Returns or sets whether the specified slicer can be modified when the sheet that contains it is protected. Read/write.
- `SlicerCache As SlicerCache  (read-only)`  
  Returns the SlicerCache object associated with the slicer. Read-only.
- `SlicerCacheLevel As SlicerCacheLevel  (read-only)`  
  Returns the SlicerCacheLevel object associated with the slicer. Read-only.
- `Shape As Shape  (read-only)`  
  Returns the Shape object associated with the specified slicer. Read-only.
- `Style As Variant  (read/write)`  
  Returns or sets the style currently applied to the specified slicer. Read/write.
- `ActiveItem As SlicerItem  (read-only)`  
  Returns a SlicerItem object that represents the slicer button that is currently in focus for the specified slicer. Read-only.
- `TimelineViewState As TimelineViewState  (read-only)`  
  The timeline-specific state of the slicer. Read-only TimelineViewState.
- `SlicerCacheType As XlSlicerCacheType  (read-only)`  
  Returns the type of the slicer cache: slicer (xlSlicer) or timeline (xlTimeline) . Read-only XlSlicerCacheType.

## Methods (3)

- `Delete()`  
  Deletes the slicer and removes it from the associated Slicers collection.
- `Cut()`  
  Cuts the specified slicer and copies it to the clipboard.
- `Copy()`  
  Copies the specified slicer to the clipboard.
