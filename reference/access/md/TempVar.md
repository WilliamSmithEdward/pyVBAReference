# TempVar

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {12DCE806-EA8A-46AA-88DF-C4486EDB78E3}  

Represents a variable that can be used in Visual Basic for Applications (VBA) code or from a macro.

**Remarks:** TempVar objects provide a convenient way to exchange data between VBA procedures and macros. Although a TempVar object can be used to store information for use in VBA procedures, it does not have the same functionality as a VBA variable. - By default, a TempVar object remains in memory until the Database is closed. Use the Remove method or the RemoveTempVar macro action to remove a TempVar object. - In VBA, a TempVar object is accessible only to the members of the Access Application object, referenced databases, or add-ins. - A TempVar object can store only text or numeric data. TempVar objects cannot store objects. To refer to a TempVar object in a collection by its ordinal number or by its Name property setting, use the following syntax form: - TempVar![name]

## Properties (2)

- `Name As String  (read-only)`  
  Gets the name of the specified TempVar object. Read-only String.
- `Value As Variant  (read/write)`  
  Gets or sets the value of the specified TempVar object. Read/write Variant.
