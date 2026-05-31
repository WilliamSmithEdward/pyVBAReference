# ControlFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024440-0000-0000-C000-000000000046}  

Contains Microsoft Excel control properties.

**Example:**

```vba
Worksheets(1).Shapes(1).ControlFormat.ListFillRange = "A1:A10"
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `DropDownLines As Long  (read/write)`  
  Returns or sets the number of list lines displayed in the drop-down portion of a combo box. Read/write Long.
- `Enabled As Boolean  (read/write)`  
  True if the object is enabled. Read/write Boolean.
- `LargeChange As Long  (read/write)`  
  Returns or sets the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region). Read/write Long.
- `LinkedCell As String  (read/write)`  
  Returns or sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell. Read/write String.
- `ListCount As Long  (read/write)`  
  Returns the number of entries in a list box or combo box. Returns 0 (zero) if there are no entries in the list. Read-only Long.
- `ListFillRange As String  (read/write)`  
  Returns or sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box. Read/write String.
- `ListIndex As Long  (read/write)`  
  Returns or sets the index number of the currently selected item in a list box or combo box. Read/write Long.
- `LockedText As Boolean  (read/write)`  
  True if the text in the specified object will be locked to prevent changes when the workbook is protected. Read/write Boolean.
- `Max As Long  (read/write)`  
  Returns or sets the maximum value of a scroll bar or spinner range. The scroll bar or spinner won't take on values greater than this maximum value. Read/write Long.
- `Min As Long  (read/write)`  
  Returns or sets the minimum value of a scroll bar or spinner range. The scroll bar or spinner won't take on values less than this minimum value. Read/write Long.
- `MultiSelect As Long  (read/write)`  
  Returns or sets the selection mode of the specified list box. Can be one of the following constants: xlNone, xlSimple, or xlExtended. Read/write Long.
- `PrintObject As Boolean  (read/write)`  
  True if the object will be printed when the document is printed. Read/write Boolean.
- `SmallChange As Long  (read/write)`  
  Returns or sets the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user chooses an arrow). Read/write Long.
- `_Default As Long  (read/write)`
- `Value As Long  (read/write)`  
  Returns or sets a Long value that represents the name of the specified control format.

## Methods (4)

- `AddItem(Text As String, [Index As Variant])`  
  Adds an item to a list box or a combo box.
    - `Text As String` (required): The text to be added.
    - `Index As Variant` (optional): The position of the new entry. If the list has fewer entries than the specified index, blank items from the end of the list are added to the specified position. If this argument is omitted, the item is appended to the existing list.
- `RemoveAllItems()`  
  Removes all entries from a Microsoft Excel list box or combo box.
- `RemoveItem(Index As Long, [Count As Variant])`  
  Removes one or more items from a list box or combo box.
    - `Index As Long` (required): The number of the first item to be removed. Valid values are from 1 to the number of items in the list (returned by the ListCount property).
    - `Count As Variant` (optional): The number of items to be removed, starting at item _Index_. If this argument is omitted, one item is removed. If _Index_ + _Count_ exceeds the number of items in the list, all items from _Index_ through the end of the list are removed without an error.
- `List([Index As Variant]) As Variant`  
  Returns or sets the text entries in the specified list box or combo box, as an array of strings, or returns or sets a single text entry. An error occurs if there are no entries in the list.
    - `Index As Variant` (optional): The index number of a single text entry to be set or returned. If this argument is omitted, the entire list is returned or set as an array of strings.
