# Windows

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020892-0000-0000-C000-000000000046}  

A collection of all the Window objects in Microsoft Excel.

**Remarks:** The Windows collection for the Application object contains all the windows in the application, whereas the Windows collection for the Workbook object contains only the windows in the specified workbook.

**Example:**

```vba
Windows.Arrange arrangeStyle:=xlCascade
```

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Window  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Window  (read-only)`
- `SyncScrollingSideBySide As Boolean  (read/write)`  
  True enables scrolling the contents of windows at the same time when documents are being compared side by side. False disables scrolling the windows at the same time.

## Methods (4)

- `Arrange([ArrangeStyle As XlArrangeStyle], [ActiveWorkbook As Variant], [SyncHorizontal As Variant], [SyncVertical As Variant]) As Variant`  
  Arranges the windows on the screen.
    - `ArrangeStyle As XlArrangeStyle` (optional): One of the constants of XlArrangeStyle specifying how the windows are arranged.
    - `ActiveWorkbook As Variant` (optional): True to arrange only the visible windows of the active workbook. False to arrange all windows. The default value is False.
    - `SyncHorizontal As Variant` (optional): Ignored if _ActiveWorkbook_ is False or omitted. True to synchronize the windows of the active workbook when scrolling horizontally. False to not synchronize the windows. The default value is False.
    - `SyncVertical As Variant` (optional): Ignored if _ActiveWorkbook_ is False or omitted. True to synchronize the windows of the active workbook when scrolling vertically. False to not synchronize the windows. The default value is False.
- `CompareSideBySideWith(WindowName As Variant) As Boolean`  
  Opens two windows in side-by-side mode. Returns a Boolean value.
    - `WindowName As Variant` (required): The name of the window.
- `BreakSideBySide() As Boolean`  
  Ends side-by-side mode if two windows are in side-by-side mode. Returns a Boolean value that represents whether the method was successful.
- `ResetPositionsSideBySide()`  
  Resets the position of two worksheet windows that are being compared side by side.
