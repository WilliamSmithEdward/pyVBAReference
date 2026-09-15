# CustomLabel

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020915-0000-0000-C000-000000000046}  

Represents a custom mailing label. The CustomLabel object is a member of the CustomLabels collection. The CustomLabels collection contains all the custom mailing labels listed in the Label Options dialog box.

**Remarks:** Use CustomLabels (Index), where Index is the custom label name or index number, to return a single CustomLabel object. The following example creates a new document with an existing custom label layout named "My Labels." The index number represents the position of the custom mailing label in the CustomLabels collection. The following example displays the name of the first custom mailing label. Use the Add method to create a custom label. The following example adds a custom mailing label named "My Label" and sets the page size.

## Properties (16)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CustomLabel object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write CustomLabel.
- `TopMargin As Single  (read/write)`  
  Returns or sets the distance (in points) between the top edge of the page and the top boundary of the body text. Read/write Single.
- `SideMargin As Single  (read/write)`  
  Returns or sets the side margin widths (in points) for the specified custom mailing label. Read/write Single.
- `Height As Single  (read/write)`  
  Returns or sets the height of a specified custom mailing label, in points. Read/write Single.
- `Width As Single  (read/write)`  
  Returns or sets the width of a custom mailing label, in points. Read/write Long.
- `VerticalPitch As Single  (read/write)`  
  Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label. Read/write Single.
- `HorizontalPitch As Single  (read/write)`  
  Returns or sets the horizontal distance (in points) between the left edge of one custom mailing label and the left edge of the next mailing label. Read/write Single.
- `NumberAcross As Long  (read/write)`  
  Returns or sets the number of custom mailing labels across a page. Read/write Long.
- `NumberDown As Long  (read/write)`  
  Returns or sets the number of custom mailing labels down the length of a page. Read/write Long.
- `DotMatrix As Boolean  (read-only)`  
  True if the printer type for the specified custom label is dot matrix. False if the printer type is either laser or ink jet. Read-only Boolean.
- `PageSize As WdCustomLabelPageSize  (read/write)`  
  Returns or sets the page size for the specified custom mailing label. Read/write WdCustomLabelPageSize.
- `Valid As Boolean  (read-only)`  
  True if the various properties (for example, Height, Width, and NumberDown) for the specified custom label work together to produce a valid mailing label. Read-only Boolean.

## Methods (1)

- `Delete()`  
  Deletes the specified custom label.
