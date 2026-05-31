# CubeField

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002444C-0000-0000-C000-000000000046}  

Represents a hierarchy or measure field from an OLAP cube. In a PivotTable report, the CubeField object is a member of the CubeFields collection.

**Example:**

```vba
Set objNewSheet = Worksheets.Add
objNewSheet.Activate
intRow = 1
For Each objPF in _
 Worksheets("Sheet1").PivotTables(1).PivotFields
 If objPF.CubeField.CubeFieldType = xlHierarchy Then
 objNewSheet.Cells(intRow, 1).Value = objPF.Name
 intRow = intRow + 1
 End If
Next objPF
```

## Properties (28)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `CubeFieldType As XlCubeFieldType  (read-only)`  
  Indicates whether the OLAP cube field is a hierarchy field or a measure field. Can be one of the XlCubeFieldType constants.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Value As String  (read-only)`  
  Returns a String value that represents the name of the specified field.
- `Orientation As XlPivotFieldOrientation  (read/write)`  
  Returns or sets an XlPivotFieldOrientation value that represents the location of the field in the specified PivotTable report.
- `Position As Long  (read/write)`  
  Returns or sets a Long value that represents the position of the hierarchy field on the PivotTable report when it's dragged from the field well.
- `TreeviewControl As TreeviewControl  (read-only)`  
  Returns the TreeviewControl object of the CubeField object, representing the cube manipulation control of an OLAP-based PivotTable report. Read-only.
- `DragToColumn As Boolean  (read/write)`  
  True if the specified field can be dragged to the column position. The default value is True. Read/write Boolean.
- `DragToHide As Boolean  (read/write)`  
  True if the field can be hidden by being dragged off the PivotTable report. The default value is True. Read/write Boolean.
- `DragToPage As Boolean  (read/write)`  
  True if the field can be dragged to the page position. The default value is True. Read/write Boolean.
- `DragToRow As Boolean  (read/write)`  
  True if the field can be dragged to the row position. The default value is True. Read/write Boolean.
- `DragToData As Boolean  (read/write)`  
  True if the specified field can be dragged to the data position. The default value is True. Read/write Boolean.
- `HasMemberProperties As Boolean  (read-only)`  
  Returns True when there are member properties specified to be displayed for the cube field. Read-only Boolean.
- `LayoutForm As XlLayoutFormType  (read/write)`  
  Returns or sets the way the specified PivotTable items appear-in table format or in outline format. Read/write XlLayoutFormType.
- `PivotFields As PivotFields  (read-only)`  
  Returns the PivotFields collection. This collection contains all PivotTable fields, including those that aren't currently visible on-screen. Read-only PivotFields object.
- `EnableMultiplePageItems As Boolean  (read/write)`  
  Set to True to allow multiple items in the page field area for OLAP PivotTables to be selected. The default value is False. Read/write Boolean.
- `LayoutSubtotalLocation As XlSubtototalLocationType  (read/write)`  
  Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field. Read/write XlSubtotalLocationType.
- `ShowInFieldList As Boolean  (read/write)`  
  When set to True (default), a CubeField object will be shown in the field list. Read/write Boolean.
- `IncludeNewItemsInFilter As Boolean  (read/write)`  
  The IncludeNewItemsInFilter property is used to track included/excluded items in OLAP PivotTables. Read/write.
- `CubeFieldSubType As XlCubeFieldSubType  (read-only)`  
  Specifies the type of a CubeField. Read-only.
- `AllItemsVisible As Boolean  (read-only)`  
  The AllItemsVisible property checks whether manual filtering is applied to a PivotField or CubeField. Read-only Boolean.
- `CurrentPageName As String  (read/write)`  
  Returns or sets the page name for a CubeField. Read/write String.
- `IsDate As Boolean  (read-only)`  
  Returns True if the CubeField is a date. Read-only Boolean.
- `Caption As String  (read/write)`  
  Returns a String value that represents the label text for the cube field.
- `FlattenHierarchies As Boolean  (read/write)`  
  Returns or sets whether items from all levels of hierarchies in a named set cube field are displayed in the same field of a PivotTable report based on an OLAP cube. Read/write.
- `HierarchizeDistinct As Boolean  (read/write)`  
  Returns or sets whether to order and remove duplicates when displaying the specified named set in a PivotTable report based on an OLAP cube. Read/write.

## Methods (5)

- `Delete()`  
  Deletes the object.
- `AddMemberPropertyField(Property As String, [PropertyOrder As Variant], [PropertyDisplayedIn As Variant])`  
  Adds a member property field to the display for the cube field.
    - `Property As String` (required): The unique name of the member property. For balanced hierarchies, a unique name can be created by appending the "quoted" member property name to the unique name of the level with which the member property is associated. For unbalanced hierarchies, a unique name can be created by appending the "quoted" member property name to the unique name of the hierarchy.
    - `PropertyOrder As Variant` (optional): Sets the PropertyOrder property value for a CubeField object. The actual position in the collection will be immediately before the PivotTable field that currently has the same PropertyOrder value that is given in the argument. If no field has the given PropertyOrder value, the range of acceptable values is 1 to the number of member properties already showing for the hierarchy plus one. This argument is one-based. If omitted, the property goes to the end of the list.
    - `PropertyDisplayedIn As Variant` (optional): Specifies where to display the property. If this argument is omitted, the member property field will be added to the PivotTable only.
- `ClearManualFilter()`  
  The ClearManualFilter method provides an easy way to set the Visible property to True for all items of a PivotField in PivotTables, and to empty the HiddenItemsList/VisibleItemsList collections in OLAP PivotTables.
- `CreatePivotFields()`  
  The CreatePivotFields method enables users to apply a filter to PivotFields not yet added to the PivotTable by creating the corresponding PivotField object.
- `AutoGroup([Orientation As Variant], [Position As Variant])`  
  Automatically groups the cube fields in an OLAP cube, optionally in the specified orientation and/or at the specified position.
    - `Orientation As Variant` (optional): The orientation of the group.
    - `Position As Variant` (optional): The position of the group.
