# OMathMat

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {3E061A7E-67AD-4EAA-BC1E-55057D5E596F}  

Represents an equation matrix.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathMat object.
- `Rows As OMathMatRows  (read-only)`  
  Returns an OMathMatRows object that represents the rows in a matrix. Read-only.
- `Cols As OMathMatCols  (read-only)`  
  Returns an OMathMatCols collection that represents the columns in a matrix. Read-only.
- `Cell As OMath  (read-only)`  
  Returns an OMath object that represents a cell in a matrix. Read-only.
- `Align As WdOMathVertAlignType  (read/write)`  
  Returns or sets a WdOMathVertAlignType constant that represents the vertical alignment for a matrix. Read/write.
- `PlcHoldHidden As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether placeholders in a matrix are hidden from display. True hides placeholder text. Read/write.
- `RowSpacingRule As WdOMathSpacingRule  (read/write)`  
  Returns or sets a WdOMathSpacingRule constant that represents the spacing rule for rows in a matrix. Read/write.
- `RowSpacing As Long  (read/write)`  
  Returns or sets a Long that represents the spacing for rows in a matrix. Read/write.
- `ColSpacing As Long  (read/write)`  
  Returns or sets a Long that represents the spacing for columns in a matrix. Read/write.
- `ColGapRule As WdOMathSpacingRule  (read/write)`  
  Returns or sets a WdOMathSpacingRule constant that represents the spacing rule for the space that appears between columns in a matrix. Read/write.
- `ColGap As Long  (read/write)`  
  Returns or sets a Long that represents the spacing between columns in a matrix. Read/write.
