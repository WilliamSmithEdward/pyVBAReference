# Styles

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002092D-0000-0000-C000-000000000046}  

A collection of Style objects that represent both the built-in and user-defined styles in a document.

**Remarks:** Use the Styles property to return the Styles collection. The following example deletes all user-defined styles in the active document. Use the Add method to create a new user-defined style and add it to the Styles collection. The following example adds a new character style named "Introduction" and makes it 12-point Arial, with bold and italic formatting. The example then applies this new character style to the selection. Use Styles (Index), where Index is the style name, a WdBuiltinStyle constant or index number, to return a single Style object. You must exactly match the spelling and spacing of the style name, but not necessarily its capitalization. The following example modifies the font of the user-defined style named "Color" in the active document. The following example sets the built-in Heading 1 style to not be bold. The style index number represents the position of the style in the alphabetically sorted list of style names. Note that Styles(1) is the first style in the alphabetical list. The following example displays the base style and style name of the first style in the Styles collection. The Styles object is not available from the Template object.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Styles object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of styles in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As Style`  
  Returns an individual Style object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, [Type As Variant]) As Style`  
  Creates a new user-defined style and adds it to the Styles collection.
    - `Name As String` (required): The new style name.
    - `Type As Variant` (optional): Can be one of the WdStyleType constants. The default is paragraph style.
