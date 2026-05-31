# ITextStream

**Type:** Dispatch Interface  
**Library:** Microsoft Scripting Runtime  
**GUID:** {53BAD8C1-E718-11CF-893D-00A0C9054228}  

Scripting.TextStream Interface

## Properties (4)

- `Line As Long  (read-only)`  
  Current line number
- `Column As Long  (read-only)`  
  Current column number
- `AtEndOfStream As Boolean  (read-only)`  
  Is the current position at the end of the stream?
- `AtEndOfLine As Boolean  (read-only)`  
  Is the current position at the end of a line?

## Methods (9)

- `Read(Characters As Long) As String`  
  Read a specific number of characters into a string
- `ReadLine() As String`  
  Read an entire line into a string
- `ReadAll() As String`  
  Read the entire stream into a string
- `Write(Text As String)`  
  Write a string to the stream
- `WriteLine([Text As String])`  
  Write a string and an end of line to the stream
- `WriteBlankLines(Lines As Long)`  
  Write a number of blank lines to the stream
- `Skip(Characters As Long)`  
  Skip a specific number of characters
- `SkipLine()`  
  Skip a line
- `Close()`  
  Close a text stream
