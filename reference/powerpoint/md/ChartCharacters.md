# ChartCharacters

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A57-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents characters in an object that contains text.

**Remarks:** The ChartCharacters object lets you modify any sequence of characters contained in the full text string. Use Characters ( Start, Length ), where Start is the start character number and Length is the number of characters, to return a ChartCharacters object.

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
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Insert(String As String) As Variant`  
  Inserts a string preceding the selected characters.
    - `String As String` (required): The string to insert.
