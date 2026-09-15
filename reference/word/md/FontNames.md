# FontNames

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002096F-0000-0000-C000-000000000046}  

Represents a list of the names of all the available fonts.

**Remarks:** Use the FontNames, LandscapeFontNames, or PortraitFontNames property to return the FontNames object. The following example displays the number of portrait fonts available. This example lists all the font names in the FontNames object at the end of the active document. Use FontNames (Index), where Index is the index number, to return the name of a font. The following example displays the first font name in the FontNames object.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of font names in the collection. Read-only Long.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FontNames object.

## Methods (1)

- `Item(Index As Long) As String`  
  Returns a String that represents the name of a font.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
