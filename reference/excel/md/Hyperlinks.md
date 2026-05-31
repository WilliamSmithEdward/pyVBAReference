# Hyperlinks

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024430-0000-0000-C000-000000000046}  

Represents the collection of hyperlinks for a worksheet or range.

**Remarks:** Each hyperlink is represented by a Hyperlink object.

**Example:**

```vba
For Each h in Worksheets(1).Hyperlinks
 If Instr(h.Name, "Microsoft") <> 0 Then h.Follow
Next
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
- `Item As Hyperlink  (read-only)`  
  Returns a single object from a collection.
- `_Default As Hyperlink  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(Anchor As Object, Address As String, [SubAddress As Variant], [ScreenTip As Variant], [TextToDisplay As Variant]) As Object`  
  Adds a hyperlink to the specified range or shape.
    - `Anchor As Object` (required): The anchor for the hyperlink. Can be either a Range or Shape object.
    - `Address As String` (required): The address of the hyperlink.
    - `SubAddress As Variant` (optional): The subaddress of the hyperlink.
    - `ScreenTip As Variant` (optional): The screen tip to be displayed when the mouse pointer is paused over the hyperlink.
    - `TextToDisplay As Variant` (optional): The text to be displayed for the hyperlink.
- `Delete()`  
  Deletes the object.
