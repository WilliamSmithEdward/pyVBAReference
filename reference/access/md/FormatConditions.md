# FormatConditions

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {E27A992E-A330-11D0-81DD-00C04FC2F51B}  

The FormatConditions collection represents the collection of conditional formats for a combo box or text box control. Each format is represented by a FormatCondition object.

**Remarks:** Use the FormatConditions property of a combo box or text box in Visual Basic or in an expression to return a FormatConditions collection. Use the Add method to create a new conditional format, and use the Modify method to change an existing conditional format. Use the Modify method to change one of the formats, or the Delete method to delete a format.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As _FormatCondition  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only FormatCondition.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.

## Methods (2)

- `Add(Type As AcFormatConditionType, [Operator As AcFormatConditionOperator], [Expression1 As Variant], [Expression2 As Variant]) As _FormatCondition`  
  Use the Add method to add a conditional format as a FormatCondition object to the FormatConditions collection of a combo box or text box control.
    - `Type As AcFormatConditionType` (required): An AcFormatConditionType constant that specifies the type of format condition to be added.
    - `Operator As AcFormatConditionOperator` (optional): An AcFormatConditionOperator constant that specifies the operator. If the Type argument is acExpression, the Operator argument is ignored. If you leave this argument blank, the default constant (acBetween) is assumed.
    - `Expression1 As Variant` (optional): A value or expression associated with the first part of the conditional format. Can be a constant or a string value.
    - `Expression2 As Variant` (optional): A value or expression associated with the second part of the conditional format when the Operator argument is acBetween or acNotBetween (otherwise, this argument is ignored). Can be a constant or a string value.
- `Delete()`  
  Deletes the specified object.
