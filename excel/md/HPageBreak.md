# HPageBreak

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024401-0000-0000-C000-000000000046}  

Represents a horizontal page break.

**Remarks:** The HPageBreak object is a member of the HPageBreaks collection.

**Example:**

```vba
Set Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Worksheet  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Type As XlPageBreak  (read/write)`  
  Returns or sets an XlPageBreak value that represents the page break type.
- `Extent As XlPageBreakExtent  (read-only)`  
  Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following XlPageBreakExtent constants: xlPageBreakFull or xlPageBreakPartial. Read-only Long.
- `Location As Range  (read/write)`  
  Returns or sets the cell (a Range object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell. Read/write Range.

## Methods (2)

- `Delete()`  
  Deletes the object.
- `DragOff(Direction As XlDirection, RegionIndex As Long)`  
  Drags a page break out of the print area.
    - `Direction As XlDirection` (required): The direction in which the page break is dragged.
    - `RegionIndex As Long` (required): The print-area region index for the page break (the region where the mouse pointer is located when the mouse button is pressed if the user drags the page break). If the print area is contiguous, there's only one print region. If the print area is discontiguous, there's more than one print region.
