# VPageBreak

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024402-0000-0000-C000-000000000046}  

Represents a vertical page break.

**Remarks:** The VPageBreak object is a member of the VPageBreaks collection.

**Example:**

```vba
Dim r as Range
Set r = Worksheets(1).VPageBreaks(1).Location
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
  Returns the cell (a Range object) that defines the page-break location. Vertical page breaks are aligned with the left edge of the location cell. Read-only Range.

## Methods (2)

- `Delete()`  
  Deletes the object.
- `DragOff(Direction As XlDirection, RegionIndex As Long)`  
  Drags a page break out of the print area.
    - `Direction As XlDirection` (required): The direction in which the page break is dragged.
    - `RegionIndex As Long` (required): The print-area region index for the page break (the region where the mouse pointer is located when the mouse button is pressed if the user drags the page break). If the print area is contiguous, there's only one print region. If the print area is discontiguous, there's more than one print region.
