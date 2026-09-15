# HeadingStyle

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002098B-0000-0000-C000-000000000046}  

Represents a style used to build a table of contents or figures. The HeadingStyle object is a member of the HeadingStyles collection.

**Remarks:** Use HeadingStyles (Index), where Index is the index number, to return a single HeadingStyle object. The index number represents the position of the style in the HeadingStyles collection. The following example adds (at the beginning of the active document) a table of figures built from the Title style, and then displays the name of the first style in the HeadingStyles collection. Use the Add method to add a style to the HeadingStyles collection. The following example adds a table of contents at the beginning of the active document and then adds the Title style to the list of styles used to build a table of contents.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HeadingStyle object.
- `Style As Variant  (read/write)`  
  Returns or sets the style for a heading. Read/write Variant.
- `Level As Integer  (read/write)`  
  Returns or sets the level for the heading style in a table of contents or table of figures. Read/write Integer.

## Methods (1)

- `Delete()`  
  Deletes the specified heading style.
