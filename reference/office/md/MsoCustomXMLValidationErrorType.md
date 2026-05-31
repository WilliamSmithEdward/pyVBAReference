# MsoCustomXMLValidationErrorType

**Type:** Enumeration  
**Library:** Microsoft Office 16.0 Object Library  

Indicates how validation errors will be cleared or generated.

## Constants (3)

- `msoCustomXMLValidationErrorSchemaGenerated` = 0  
  Specifies that where there is a non-empty schema collection available for the custom XML part and validation is in effect, any changes to the part will cause validation errors.
- `msoCustomXMLValidationErrorAutomaticallyCleared` = 1  
  Specifies that the error will clear itself whenever any change is made to the node it is bound to.
- `msoCustomXMLValidationErrorManual` = 2  
  Specifies that the error will not be cleared until the Delete method is called.
