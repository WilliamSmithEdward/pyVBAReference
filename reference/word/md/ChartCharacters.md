# ChartCharacters

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {FF06FEF2-DA89-41C0-A0A8-5CD434E210AD}  

Represents characters in an object that contains text.

**Remarks:** The ChartCharacters object lets you modify any sequence of characters contained in the full text string. Use Characters ( _Start_ , _Length_ ), where _Start_ is the start character number and _Length_ is the number of characters, to return a ChartCharacters object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1).Chart
 .ChartTitle.Characters.Font.Italic = True
End With
```

## Properties (8)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Caption As String  (read/write)`  
  Returns the text of this range of characters. Read-only String.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Font As ChartFont  (read-only)`  
  Returns the font of the specified object. Read-only ChartFont.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `PhoneticCharacters As String  (read/write)`  
  Returns or sets the phonetic text for the object. Read/write String.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Insert(String As String) As Variant`  
  Inserts a string preceding the selected characters.
    - `String As String` (required): The string to insert.
