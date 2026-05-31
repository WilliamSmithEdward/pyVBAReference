# Phonetic

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024438-0000-0000-C000-000000000046}  

Contains information about a specific phonetic text string in a cell.

**Remarks:** In Microsoft Excel 97, this object contained the formatting attributes for any phonetic text in the specified range.

**Example:**

```vba
ActiveCell.Phonetics(1).Text = ""
```

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `CharacterType As Long  (read/write)`  
  Returns or sets the type of phonetic text in the specified cell. Read/write XlPhoneticCharacterType.
- `Alignment As Long  (read/write)`  
  Returns or sets a Long value that represents the alignment for the specified phonetic text or tick label.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
