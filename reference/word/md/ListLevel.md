# ListLevel

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002098D-0000-0000-C000-000000000046}  

Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list. The ListLevel object is a member of the ListLevels collection.

**Remarks:** Use ListLevels (Index), where Index is a number from 1 through 9, to return a single ListLevel object. The following example sets list level one of list template one in the active document to start at 4. The ListLevel object gives you access to all the formatting properties for the specified list level, such as the Alignment, Font, NumberFormat, NumberPosition, NumberStyle, and TrailingCharacter properties. To apply a list level, first identify the range or list, and then use the ApplyListTemplate method. Each tab at the beginning of the paragraph is translated into a list level. For example, a paragraph that begins with three tabs will become a level-three list paragraph after the ApplyListTemplate method is used.

## Properties (16)

- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `NumberFormat As String  (read/write)`  
  Returns or sets the number format for the specified list level. Read/write String.
- `TrailingCharacter As WdTrailingCharacter  (read/write)`  
  Returns or sets the character inserted after the number for the specified list level. Read/write WdTrailingCharacter.
- `NumberStyle As WdListNumberStyle  (read/write)`  
  Returns or sets the number style for the ListLevel object. Read/write WdListNumberStyle.
- `NumberPosition As Single  (read/write)`  
  Returns or sets the position (in points) of the number or bullet for the specified ListLevel object. Read/write Single.
- `Alignment As WdListLevelAlignment  (read/write)`  
  Returns or sets a WdListLevelAlignment constant that represents the alignment for the list level of the list template. Read/write.
- `TextPosition As Single  (read/write)`  
  Returns or sets the position (in points) for the second line of wrapping text for the specified ListLevel object. Read/write Single.
- `TabPosition As Single  (read/write)`  
  Returns or sets the tab position for the specified ListLevel object. Read/write Single.
- `StartAt As Long  (read/write)`  
  Returns or sets the starting number for the specified ListLevel object. Read/write Long.
- `LinkedStyle As String  (read/write)`  
  Returns or sets the name of the style that's linked to the specified ListLevel object. Read/write String.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified object. Read/write Font.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListLevel object.
- `ResetOnHigher As Long  (read/write)`  
  Sets or returns the list level that must appear before the specified list level restarts numbering at 1. Read/write Long.
- `PictureBullet As InlineShape  (read-only)`  
  Returns an InlineShape object that represents a picture bullet.

## Methods (1)

- `ApplyPictureBullet(FileName As String) As InlineShape`  
  Formats a paragraph or range of paragraphs with a picture bullet.
    - `FileName As String` (required): The path and file name of the picture file.
