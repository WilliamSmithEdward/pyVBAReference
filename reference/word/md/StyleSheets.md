# StyleSheets

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {07B7CC7E-E66C-11D3-9454-00105AA31A08}  

A collection of StyleSheet objects that represents the cascading style sheets attached to a document. The StyleSheets collection includes all cascading style sheets displayed in the Linked CSS Style Sheets dialog box.

**Remarks:** Use the StyleSheets property to return the StyleSheets collection. Use the Add method to add a style sheet to the StyleSheets collection. The following example adds three cascading style sheets to the active document and sets the third as the highest in precedence.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified StyleSheets object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of style sheets in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As StyleSheet`  
  Returns an individual StyleSheet object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(FileName As String, LinkType As WdStyleSheetLinkType, Title As String, Precedence As WdStyleSheetPrecedence) As StyleSheet`  
  Returns a StyleSheet object that represents a new style sheet added to a web document.
    - `FileName As String` (required): The path and file name of the cascading style sheet.
    - `LinkType As WdStyleSheetLinkType` (required): Indicates whether the style sheet should be added as a link or imported into the web document.
    - `Title As String` (required): The name of the style sheet.
    - `Precedence As WdStyleSheetPrecedence` (required): Indicates the level of importance compared with other cascading style sheets attached to the web document.
