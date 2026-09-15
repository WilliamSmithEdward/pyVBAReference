# HeadingStyles

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002098A-0000-0000-C000-000000000046}  

A collection of HeadingStyle objects that represent the styles used to compile a table of figures or table of contents.

**Remarks:** Use the HeadingStyles property to return the HeadingStyles collection. The following example displays the number of items in the HeadingStyles collection for the first table of contents in the active document. Use the Add method to add a style to the HeadingStyles collection. The following example adds a table of contents at the beginning of the active document and then adds the Title style to the list of styles used to build a table of contents. Use HeadingStyles (Index), where Index is the index number, to return a single HeadingStyle object. The index number represents the position of the style in the HeadingStyles collection. The following example adds (at the beginning of the active document) a table of figures built from the Title style, and then displays the name of the first style in the HeadingStyles collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HeadingStyles object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of heading styles in the collection. Read-only.

## Methods (2)

- `Item(Index As Long) As HeadingStyle`  
  Returns an individual HeadingStyle object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Style As Variant, Level As Integer) As HeadingStyle`  
  Returns a HeadingStyle object that represents a new heading style added to a document. The new heading style will be included whenever you compile a table of contents or table of figures.
    - `Style As Variant` (required): The style you want to add. You can specify this argument by using either the string name for the style or a Style object.
    - `Level As Integer` (required): A number that represents the level of the heading.
