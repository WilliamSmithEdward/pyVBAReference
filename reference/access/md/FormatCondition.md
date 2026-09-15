# FormatCondition

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {E27A992D-A330-11D0-81DD-00C04FC2F51B}  

The FormatCondition object represents a conditional format of a combo box or text box control and is a member of the FormatConditions collection.

**Remarks:** Use the FormatConditions (_index_), where _index_ is the index number of the conditional format, to return a FormatCondition object. Use the Add method to create a new conditional format. Use the Modify method to change one of the formats, or the Delete method to delete a format. Use the BackColor, Enabled, FontBold, FontItalic, FontUnderline, and ForeColor properties of the FormatCondition object to control the appearance of formatted combo box and text box controls.

## Properties (15)

- `ForeColor As Long  (read/write)`  
  Use the ForeColor property to specify the color for text in a control. Read/write Long.
- `BackColor As Long  (read/write)`  
  Gets or sets the interior color of the specified object. Read/write Long.
- `FontBold As Boolean  (read/write)`  
  Use the FontBold property to specify whether a font appears in a bold style in the following situations:
- `FontItalic As Boolean  (read/write)`  
  Use the FontItalic property to specify whether text is italic in the following situations:
- `FontUnderline As Boolean  (read/write)`  
  Use the FontUnderline property to specify whether text is underlined in the following situations:
- `Enabled As Boolean  (read/write)`  
  Use the Enabled property to set or return the status of the conditional format in the FormatCondition object. Read/write Boolean.
- `Type As AcFormatConditionType  (read-only)`  
  Returns the value of a FormatCondition object type. Read-only AcFormatConditionType.
- `Operator As AcFormatConditionOperator  (read-only)`  
  Use the Operator property to return the operator value for the conditional format or data validation of a FormatCondition object. Read-only AcFormatConditionOperator.
- `Expression1 As String  (read-only)`  
  Use the Expression1 property to return the values of a conditional format within a FormatCondition object. Read-only String.
- `Expression2 As String  (read-only)`  
  Use the Expression2 property to return the values of a conditional format within a FormatCondition object. Read-only String.
- `ShortestBarLimit As AcFormatBarLimits  (read/write)`  
  Gets or sets how the shortest data bar is evaluated for the specified FormatCondition. Read/write AcFormatBarLimits.
- `ShortestBarValue As String  (read/write)`  
  Gets or sets a numeric expression that specifies the value of the shortest bar of a FormatCondition. Read/write String.
- `LongestBarLimit As AcFormatBarLimits  (read/write)`  
  Gets or sets how the longest data bar is evaluated for the specified FormatCondition. Read/write AcFormatBarLimits.
- `LongestBarValue As String  (read/write)`  
  Gets or sets a numeric expression that specifies the value of the longest bar of a FormatCondition. Read/write String.
- `ShowBarOnly As Boolean  (read/write)`  
  Gets or sets whether the data bar or the data bar and its value are displayed. Set to True to display only the data bar. Read/write Boolean.

## Methods (2)

- `Modify(Type As AcFormatConditionType, [Operator As AcFormatConditionOperator], [Expression1 As Variant], [Expression2 As Variant])`  
  Use the Modify method to change the format conditions of a FormatCondition object in the FormatConditions collection of a combo box or text box control.
    - `Type As AcFormatConditionType` (required): An AcFormatConditionType constant that specifies the type of condition to be modified.
    - `Operator As AcFormatConditionOperator` (optional): An AcFormatConditionOperator constant that specifies the type of operator to be used. NOTE: If the type argument is acExpression, the operator argument is ignored. If you leave this argument blank, the default constant (acBetween) is assumed.
    - `Expression1 As Variant` (optional): A value or expression associated with the first part of the conditional format. Can be a constant value or a string value.
    - `Expression2 As Variant` (optional): A value or expression associated with the second part of the conditional format when the operator argument is acBetween or acNotBetween (otherwise, this argument is ignored). Can be a constant value or a string value.
- `Delete()`  
  Deletes the specified object.
