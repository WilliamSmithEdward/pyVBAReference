# Module

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {331FDCFE-CF31-11CD-8701-00AA003F0F07}  

A Module object refers to a standard module or a class module.

**Remarks:** Microsoft Access includes class modules that are not associated with any object, and form modules and report modules, which are associated with a form or report. To determine whether a Module object represents a standard module or a class module from code, check the Module object's Type property. The Modules collection contains all open Module objects, regardless of their type. Modules in the Modules collection can be compiled or uncompiled. To return a reference to a particular standard or class Module object in the Modules collection, use any of the following syntax forms. The following example returns a reference to a standard Module object and assigns it to an object variable. Note that the brackets enclosing the name of the Module object are necessary only if the name of the Module object includes spaces. The next example returns a reference to a form Module object and assigns it to an object variable. To refer to a specific form or report module, you can also use the Form or Report object's Module property. The following example also returns a reference to the Module object associated with an Employees form and assigns it to an object variable.

## Properties (11)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Lines As String  (read-only)`  
  The Lines property returns a string containing the contents of a specified line or lines in a standard module or a class module. Read-only String.
- `CountOfLines As Long  (read-only)`  
  The CountOfLines property returns a Long value indicating the number of lines of code in a standard module or class module. Read-only Long.
- `ProcStartLine As Long  (read-only)`  
  The ProcStartLine property returns a value identifying the line at which a specified procedure begins in a standard module or a class module. Read-only Long.
- `ProcCountLines As Long  (read-only)`  
  The ProcCountLines property returns the number of lines in a specified procedure in a standard module or a class module. Read-only Long.
- `ProcBodyLine As Long  (read-only)`  
  The ProcBodyLine property returns the number of the line at which the body of a specified procedure begins in a standard module or a class module. Read-only Long.
- `ProcOfLine As String  (read-only)`  
  The ProcOfLine property returns the name of the procedure that contains a specified line in a standard module or a class module. Read-only string.
- `CountOfDeclarationLines As Long  (read-only)`  
  The CountOfDeclarationLines property returns a Long value indicating the number of lines of code in the Declarations section in a standard module or class module. Read-only Long.
- `Type As AcModuleType  (read-only)`  
  Indicates whether a module is a standard module or a class module. Read-only AcModuleType.

## Methods (8)

- `InsertText(Text As String)`  
  The InsertText method inserts a specified string of text into a standard module or a class module.
    - `Text As String` (required): The text to be inserted into the module.
- `AddFromString(String As String)`  
  The AddFromString method adds a string to a Module object. The Module object may represent a standard module or a class module.
    - `String As String` (required): The information that you want to add to the module.
- `AddFromFile(FileName As String)`  
  The AddFromFile method adds the contents of a text file to a Module object. The Module object may represent a standard module or a class module.
    - `FileName As String` (required): The name and full path of a text (.txt) file or another file that stores text in an ANSI format.
- `InsertLines(Line As Long, String As String)`  
  The InsertLines method inserts a line or group of lines of code in a standard module or a class module.
    - `Line As Long` (required): The number of the line at which to begin inserting.
    - `String As String` (required): The text to be inserted into the module.
- `DeleteLines(StartLine As Long, Count As Long)`  
  The DeleteLines method deletes lines from a standard module or a class module.
    - `StartLine As Long` (required): The number of the line from which to begin deleting.
    - `Count As Long` (required): The number of lines to delete.
- `ReplaceLine(Line As Long, String As String)`  
  The ReplaceLine method replaces a specified line in a standard module or a class module.
    - `Line As Long` (required): The number of the line to be replaced.
    - `String As String` (required): The text that is to replace the existing line.
- `CreateEventProc(EventName As String, ObjectName As String) As Long`  
  The CreateEventProc method creates an event procedure in a class module.
    - `EventName As String` (required): The name of an event.
    - `ObjectName As String` (required): An object that has the event specified by the _EventName_ argument. If the event procedure is being added to a Form, the word "Form" should be specified for this argument. If the event procedure is being added to a Report, the word "Report" should be specified for this argument. If the event procedure is being added to a Control, the name of the control should be specified for this argument.
- `Find(Target As String, StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long, [WholeWord As Boolean], [MatchCase As Boolean], [PatternSearch As Boolean]) As Boolean`  
  Finds specified text in a standard module or class module.
    - `Target As String` (required): The text that you want to find.
    - `StartLine As Long` (required): The line on which to begin searching. If a match is found, the value of the _StartLine_ argument is set to the line on which the beginning character of the matching text is found.
    - `StartColumn As Long` (required): The column on which to begin searching. Each character in a line is in a separate column, beginning with zero on the left side of the module. If a match is found, the value of the _StartColumn_ argument is set to the column in which the beginning character of the matching text is found.
    - `EndLine As Long` (required): The line on which to stop searching. If a match is found, the value of the _EndLine_ argument is set to the line on which the ending character of the matching text is found.
    - `EndColumn As Long` (required): The column on which to stop searching. If a match is found, the value of the _EndColumn_ argument is set to the column in which the beginning character of the matching text is found.
    - `WholeWord As Boolean` (optional): True results in a search for whole words only. The default is False.
    - `MatchCase As Boolean` (optional): True results in a search for words with case matching the _Target_ argument. The default is False.
    - `PatternSearch As Boolean` (optional): True results in a search in which the _Target_ argument may contain wildcard characters such as an asterisk () or a question mark (?). The default is False**.
