# TextColumn2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03B2-0000-0000-C000-000000000046}  

Represents a single text column. The TextColumn2 object is a member of the TextColumns2 collection.

**Remarks:** Use TextColumns2(_index_), where _index_ is the index number, to return a single TextColumn2 object. The index number represents the position of the column in the TextColumns2 collection (counting from left to right).

**Example:**

```vba
ActiveDocument.PageSetup.TextColumns2.Add _
 Width:=InchesToPoints(2.5), _
 Spacing:=InchesToPoints(0.5), EvenlySpaced:=False
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the TextColumn2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the TextColumn2 object was created. Read-only.
- `Number As Long  (read/write)`  
  Gets or sets the index of the TextColumn2 object. Read/write.
- `Spacing As Single  (read/write)`  
  Gets or sets the spacing between text columns in a TextColumn2 object. Read/write.
- `TextDirection As MsoTextDirection  (read/write)`  
  Gets or sets the direction of text in the TextColumn2 object. Read/write.
