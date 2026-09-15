# OMathEqArray

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {1F998A61-71C6-44C2-A0F2-1D66169B47CB}  

Represents a mathematical equation array object, consisting of one or more equations that can be vertically justified as a unit respect to surrounding text on the line.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathEqArray object.
- `E As OMathArgs  (read-only)`  
  Returns an OMathArgs object that represents the list of arguments for the specified equation object. Read-only.
- `MaxDist As Boolean  (read/write)`  
  Returns or sets a Boolean that represents that the equations in an equation array are distributed equally within the margins of its container, such as a column, cell, or page width. Read/write.
- `ObjDist As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the equations in an equation array are spaced to the maximum width of the equation array. Read/write.
- `Align As WdOMathVertAlignType  (read/write)`  
  Returns or sets a WdOMathVertAlignType that represents the type of vertical alignment for an equation array with respect to the text that surrounds the array. Read/write.
- `RowSpacingRule As WdOMathSpacingRule  (read/write)`  
  Returns or sets a WdOMathSpacingRule that represents the rule that defines spacing in an equation array. Read/write.
- `RowSpacing As Long  (read/write)`  
  Returns or sets a Long that represents the spacing between the rows in an equation array. Read/write.
