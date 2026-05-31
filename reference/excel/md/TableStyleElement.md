# TableStyleElement

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A5-0000-0000-C000-000000000046}  

Represents a single table style element.

**Remarks:** Table styles define formatting for one or all of the elements of a table, PivotTable, or slicer. For example, a header row is an element of a table. A table style can define that the fill color of the header row is red. Each table style element in a table can have formatting specified in a table style applied to the element.

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `HasFormat As Boolean  (read-only)`  
  Returns whether a table style element has formatting applied to the specified element. Read-only Boolean.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that represents the interior of the specified object. Read-only.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that represents the borders of a TableStyleElement object. Read-only.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object. Read-only.
- `StripeSize As Long  (read/write)`  
  Returns or sets the size of banding. Read/write Long.

## Methods (1)

- `Clear()`  
  Clears the formatting for this element.
