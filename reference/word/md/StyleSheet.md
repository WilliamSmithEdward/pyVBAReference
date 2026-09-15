# StyleSheet

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209EF-0000-0000-C000-000000000046}  

Represents a single cascading style sheet attached to a web document. The StyleSheet object is a member of the StyleSheets collection. The StyleSheets collection contains all the cascading style sheets attached to a specified document.

**Remarks:** Use the Item method or StyleSheets (Index), where Index is the name or number of the style sheet, of the StyleSheets collection to return a StyleSheet object. The following example removes the second style sheet from the StyleSheets collection. Use the Index property to determine the precedence of cascading style sheets. The following example creates a table of attached cascading style sheets, ordered and indexed according to which style sheet is most important. Use the Move method to reorder the precedence of attached style sheets. The following example moves the most important style sheet to the least important of all attached cascading style sheets.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified StyleSheet object.
- `FullName As String  (read-only)`  
  Specifies the name of a cascading style sheet, including the drive or Web path. Read-only String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Path As String  (read-only)`  
  Returns the disk or Web path to the specified style sheet. Read-only String.
- `Type As WdStyleSheetLinkType  (read/write)`  
  Returns or sets the style sheet type. Read/write WdStyleSheetLinkType.
- `Title As String  (read/write)`  
  Returns or sets a String representing the title of a Web style sheet. Read/write.

## Methods (2)

- `Move(Precedence As WdStyleSheetPrecedence)`  
  Moves a style sheet's order of precedence.
    - `Precedence As WdStyleSheetPrecedence` (required): The precedence level.
- `Delete()`  
  Deletes the specified cascading style sheet.
