# OMath

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E4442A83-F623-459C-8E95-8BFB44DCF23A}  

Represents an equation. OMath objects are members of the OMaths collection.

**Remarks:** Use the Add method of the OMaths collection to create an equation and add it to a document, selection, or range. The following example creates an equation and uses the BuildUp method to convert the equation to professional format.

## Properties (17)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMath object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
- `Functions As OMathFunctions  (read-only)`  
  Returns an OMathFunctions collection that represents the functions contained within an equation. Read-only.
- `Type As WdOMathType  (read/write)`  
  Returns or sets a WdOMathType constant that represents whether an equation is displayed inline with the text around it or displayed on its own line. Read/write.
- `ParentOMath As OMath  (read-only)`  
  Returns an OMath object that represents the parent element. Read-only.
- `ParentFunction As OMathFunction  (read-only)`  
  Returns an OMathFunction object that represents the parent, or containing, function. Read-only.
- `ParentRow As OMathMatRow  (read-only)`  
  Returns an OMathMatRow object that represents the parent row in a matrix. Read-only.
- `ParentCol As OMathMatCol  (read-only)`  
  Returns an OMathMatCol object that represents the parent column in a matrix. Read-only.
- `ParentArg As OMath  (read-only)`  
  Returns an OMath object that represents the parent, or containing, argument. Read-only.
- `ArgIndex As Long  (read-only)`  
  Returns a Long that represents the argument index of this component relative to the containing math object. Read-only.
- `NestingLevel As Long  (read-only)`  
  Returns a Long that represents the nesting level for an OMath object. Read-only.
- `ArgSize As Long  (read/write)`  
  Returns or sets a Long that represents the script size of an argument, for example, text, script, or script-script. Read/write.
- `Breaks As OMathBreaks  (read-only)`  
  Returns an OMathBreaks collection that represents the line breaks in an equation. Read-only.
- `Justification As WdOMathJc  (read/write)`  
  Returns or sets a WdOMathJc that represents the justification for an equation. Read/write.
- `AlignPoint As Long  (read/write)`  
  Returns or sets a Long that represents the character position of the alignment point in the equation. Read/write.

## Methods (6)

- `Linearize()`  
  Converts an equation to linear format.
- `BuildUp()`  
  Converts an equation to professional format.
- `Remove()`  
  Removes an equation from the collection of equations in a document, range, or selection.
- `ConvertToMathText()`  
  Converts an equation to math text.
- `ConvertToNormalText()`  
  Converts an equation to normal text.
- `ConvertToLiteralText()`  
  Converts an equation to literal text.
