# Operation

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {96EFA5B6-F286-4590-96B5-F944707646A1}  

Represents an operation defined for an Entity object.

**Remarks:** Use the WSParameters property to return the parameters defined for the specified operation. Use the Execute method to execute the operation.

## Properties (3)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `WSParameters As WSParameters  (read-only)`  
  Gets the collection of parameters defined for the specified operation. Read-only WSParameters.

## Methods (1)

- `Execute([bstrParameters As String]) As Variant`  
  Executes the specified operation.
    - `bstrParameters As String` (optional): Specifies values for the parameters of the operation.
