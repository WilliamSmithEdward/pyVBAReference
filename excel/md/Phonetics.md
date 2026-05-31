# Phonetics

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024447-0000-0000-C000-000000000046}  

A collection of all the Phonetic objects in the specified range.

**Remarks:** Each Phonetic object contains information about a specific phonetic text string.

**Example:**

```vba
Range("A1:C4").Phonetics.Visible = True
```

## Properties (14)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Start As Long  (read-only)`  
  Returns the position that represents the first character of a phonetic text string in the specified cell. Read-only Long.
- `Length As Long  (read-only)`  
  Returns a Long value that represents the number of characters of phonetic text from the position you've specified with the Start property.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `CharacterType As Long  (read/write)`  
  Returns or sets the type of phonetic text in the specified cell. Read/write XlPhoneticCharacterType.
- `Alignment As Long  (read/write)`  
  Returns or sets a Long value that represents the alignment for the specified phonetic text or tick label.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Item As Object  (read-only)`  
  Returns a single object from a collection.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `_Default As Object  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Delete()`  
  Deletes the object.
- `Add(Start As Long, Length As Long, Text As String)`  
  Adds phonetic text to the specified cell.
    - `Start As Long` (required): The position that represents the first character in the specified cell.
    - `Length As Long` (required): The number of characters from the _Start_ position to the end of the text in the cell.
    - `Text As String` (required): Collectively, the characters that represent the phonetic text in the cell.
