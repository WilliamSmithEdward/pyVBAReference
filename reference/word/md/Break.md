# Break

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {79635BF1-BD1D-4B3F-A520-C1106F1AAAD8}  

Represents individual page, column, and section breaks in a page. Use the Break object and the related methods and properties for programmatically defining page layout in a document.

**Remarks:** Use the Item method of the Breaks collection to return a specific Break object. The following example returns the first break in the first page of the active document.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Break object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained in the specified object.
- `PageIndex As Long  (read-only)`  
  Returns a Long that represents the page number on which the specified break occurs.
