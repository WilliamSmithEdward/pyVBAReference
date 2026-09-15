# OMathDelim

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {C94688A6-A2A7-4133-A26D-726CD569D5F3}  

Represents a delimiter object, consisting of opening and closing delimiters (such as parentheses, braces, brackets, or vertical bars), and one or more elements contained inside the delimiters.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathDelim object.
- `E As OMathArgs  (read-only)`  
  Returns an OMathArgs object that represents the list of arguments for the specified equation object. Read-only.
- `BegChar As Integer  (read/write)`  
  Returns or sets an Integer that represents the beginning delimiter character in a delimiter object. Read/write.
- `SepChar As Integer  (read/write)`  
  Returns or sets an Integer that represents the separator character in a delimiter object when the delimiter object contains two or more arguments. Read/write.
- `EndChar As Integer  (read/write)`  
  Returns or sets an Integer that represents the ending delimiter character in a delimiter object. Read/write.
- `Grow As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether delimiter characters grow to the full height of the arguments that they contain. Read/write.
- `Shape As WdOMathShapeType  (read/write)`  
  Returns or sets a WdOMathShapeType constant that represents the appearance of delimiters (for example, parentheses, braces, and brackets) in relationship to the content that they surround. Read/write.
- `NoLeftChar As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the opening delimiter in a delimiter object. Read/write.
- `NoRightChar As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to hide the closing delimiter in a delimiter object. Read/write.
