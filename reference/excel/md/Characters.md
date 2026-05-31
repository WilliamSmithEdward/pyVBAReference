# Characters

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020878-0000-0000-C000-000000000046}  

Represents characters in an object that contains text.

**Remarks:** The Characters object lets you modify any sequence of characters contained in the full text string. Use Characters (_start_, _length_), where _start_ is the start character number and _length_ is the number of characters, to return a Characters object.

**Example:**

```vba
With Worksheets("Sheet1").Range("B1")
 .Value = "New Title"
 .Characters(5, 5).Font.Bold = True
End With
```

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Caption As String  (read/write)`  
  Returns a String value that represents the text of this range of characters.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `PhoneticCharacters As String  (read/write)`  
  Returns or sets the phonetic text in the specified Characters object. Read/write String.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Insert(String As String) As Variant`  
  Inserts a string preceding the selected characters.
    - `String As String` (required): The string to insert.
