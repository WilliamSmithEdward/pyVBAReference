# Name

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208B9-0000-0000-C000-000000000046}  

Represents a defined name for a range of cells. Names can be either built-in names-such as Database, Print_Area, and Auto_Open-or custom names.

**Remarks:** The Name object is a member of the Names collection for the Application, Workbook, and Worksheet objects. Use Names (_index_), where _index_ is the name index number or defined name, to return a single Name object. The index number indicates the position of the name within the collection. Names are placed in alphabetic order, from a to z, and are not case-sensitive. Although a Range object can have more than one name, there's no Names collection for the Range object. Use Name with a Range object to return the first name from the list of names (sorted alphabetically) assigned to the range.

**Example:**

```vba
MsgBox Names(1).RefersTo
```

## Properties (21)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Category As String  (read/write)`  
  Returns or sets the category for the specified name in the language of the macro. The name must refer to a custom function or command. Read/write String.
- `CategoryLocal As String  (read/write)`  
  Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command. Read/write String.
- `MacroType As XlXLMMacroType  (read/write)`  
  Returns or sets what the name refers to. Read/write XlXLMMacroType.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `RefersTo As Variant  (read/write)`  
  Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign. Read/write String.
- `ShortcutKey As String  (read/write)`  
  Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command. Read/write String.
- `Value As String  (read/write)`  
  Returns or sets a String value that represents the formula that the name is defined to refer to.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `NameLocal As String  (read/write)`  
  Returns or sets the name of the object, in the language of the user. Read/write String.
- `RefersToLocal As Variant  (read/write)`  
  Returns or sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign. Read/write String.
- `RefersToR1C1 As Variant  (read/write)`  
  Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign. Read/write String.
- `RefersToR1C1Local As Variant  (read/write)`  
  Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign. Read/write String.
- `RefersToRange As Range  (read-only)`  
  Returns the Range object referred to by a Name object. Read-only.
- `Comment As String  (read/write)`  
  Returns or sets the comment associated with the name. Read/write String.
- `WorkbookParameter As Boolean  (read/write)`  
  Returns or sets the specified Name object as a workbook parameter. Read/write.
- `ValidWorkbookParameter As Boolean  (read-only)`  
  Returns True if the specified Name object is a valid workbook parameter. Read-only Boolean.

## Methods (1)

- `Delete()`  
  Deletes the object.
