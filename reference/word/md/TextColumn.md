# TextColumn

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020974-0000-0000-C000-000000000046}  

Represents a single text column. The TextColumn object is a member of the TextColumns collection. The TextColumns collection includes all the columns in a document or section of a document.

**Remarks:** Use TextColumns (Index), where Index is the index number, to return a single TextColumn object. The index number represents the position of the column in the TextColumns collection (counting from left to right). The following example sets the space after the first text column in the active document to 0.5 inch. Use the Add method to add a column to the collection of columns. By default, there is one text column in the TextColumns collection. The following example adds a 2.5-inch-widecolumn to the active document. Use the SetCount method to arrange text into columns. The following example arranges the text in the active document into three columns.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TextColumn object.
- `Width As Single  (read/write)`  
  Returns or sets the width, in points, of the specified text columns. Read/write Long.
- `SpaceAfter As Single  (read/write)`  
  Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write Single.
