# Parameter

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002442A-0000-0000-C000-000000000046}  

Represents a single parameter used in a parameter query.

**Remarks:** The Parameter object is a member of the Parameters collection.

**Example:**

```vba
With Worksheets(1).QueryTables(1).Parameters(1)
 .SetParam xlPrompt, "Please " & .PromptString
End With
```

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `DataType As XlParameterDataType  (read/write)`  
  Returns or sets an XlParameterDataType value that represents the data type of the specified query parameter.
- `Type As XlParameterType  (read-only)`  
  Returns an XlParameterType value that represents the parameter type.
- `PromptString As String  (read-only)`  
  Returns the phrase that prompts the user for a parameter value in a parameter query. Read-only String.
- `Value As Variant  (read-only)`  
  Returns a Variant value that represents the parameter value.
- `SourceRange As Range  (read-only)`  
  Returns a Range object that represents the cell that contains the value of the specified query parameter. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `RefreshOnChange As Boolean  (read/write)`  
  True if the specified query table is refreshed whenever you change the parameter value of a parameter query. Read/write Boolean.

## Methods (1)

- `SetParam(Type As XlParameterType, Value As Variant)`  
  Defines a parameter for the specified query table.
    - `Type As XlParameterType` (required): One of the constants of XlParameterType, which specifies the parameter type.
    - `Value As Variant` (required): The value of the specified parameter, as shown in the description of the _Type_ argument.
