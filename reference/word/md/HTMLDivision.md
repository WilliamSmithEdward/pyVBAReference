# HTMLDivision

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E7-0000-0000-C000-000000000046}  

Represents a single HTML DIV element within a web document. The HTMLDivision object is a member of the HTMLDivisions collection.

**Remarks:** Use HTMLDivisions (Index), where Index refers to the HTML division in the document, to return a single HTMLDivision object. Use the Borders property to format border properties for an HTML division. This example formats three nested divisions in the active document. This example assumes that the active document is an HTML document with at least three divisions. HTML divisions can be nested within multiple HTML divisions. Use the HTMLDivisionParent method to access a parent HTML division of the current HTML division. This example formats the borders for two HTML divisions in the active document. This example assumes that the active document is an HTML document with at least two divisions.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HTMLDivision object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained in the specified object.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `LeftIndent As Single  (read/write)`  
  Returns or sets a Single that represents the left indent value (in points) for the specified HTML division. Read/write.
- `RightIndent As Single  (read/write)`  
  Returns or sets the right indent (in points) for the specified paragraphs. Read/write Single.
- `SpaceBefore As Single  (read/write)`  
  Returns or sets the spacing (in points) before the specified HTML DIV element. Read/write Single.
- `SpaceAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in points) after the specified HTML DIV element. Read/write Single.
- `HTMLDivisions As HTMLDivisions  (read-only)`  
  Returns an HTMLDivisions object that represents an HTML division in a web document.

## Methods (2)

- `HTMLDivisionParent([LevelsUp As Variant]) As HTMLDivision`  
  Returns an HTMLDivision object that represents a parent division of the current HTML division.
    - `LevelsUp As Variant` (optional): The number of parent divisions to count back to return the desired division. If the LevelsUp argument is omitted, the HTML division returned is one level up from the current HTML division.
- `Delete()`  
  Deletes the specified HTML division.
