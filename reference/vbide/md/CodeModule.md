# CodeModule

**Type:** Class  
**Library:** Microsoft Visual Basic for Applications Extensibility 5.3  
**GUID:** {0002E170-0000-0000-C000-000000000046}  

## Properties (10)

- `Parent As VBComponent  (read-only)`
- `VBE As VBE  (read-only)`
- `Lines As String  (read-only)`
- `CountOfLines As Long  (read-only)`
- `ProcStartLine As Long  (read-only)`
- `ProcCountLines As Long  (read-only)`
- `ProcBodyLine As Long  (read-only)`
- `ProcOfLine As String  (read-only)`
- `CountOfDeclarationLines As Long  (read-only)`
- `CodePane As CodePane  (read-only)`

## Methods (7)

- `AddFromString(String As String)`
- `AddFromFile(FileName As String)`
- `InsertLines(Line As Long, String As String)`
- `DeleteLines(StartLine As Long, [Count As Long])`
- `ReplaceLine(Line As Long, String As String)`
- `CreateEventProc(EventName As String, ObjectName As String) As Long`
- `Find(Target As String, StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long, [WholeWord As Boolean], [MatchCase As Boolean], [PatternSearch As Boolean]) As Boolean`
