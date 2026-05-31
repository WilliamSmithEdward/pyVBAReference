# MetaProperty

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C038F-0000-0000-C000-000000000046}  

Represents a single property in a collection of properties describing the metadata stored in a document.

**Example:**

```vba
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String
Dim result As String

result = metaProps(1).Validate
ValidateMetaProperty = result
End Function
```

## Properties (9)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the MetaProperty object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the MetaProperty object was created. Read-only.
- `Value As Variant  (read/write)`  
  Gets or sets the value of the MetaProperty object. Read/write.
- `Name As String  (read-only)`  
  Gets the name of the MetaProperty object. Read-only.
- `Id As String  (read-only)`  
  Gets the ID of the MetaProperty object. Read-only.
- `IsReadOnly As Boolean  (read-only)`  
  Gets a Boolean value that specifies whether the meta property is read-only. Read-only.
- `IsRequired As Boolean  (read-only)`  
  Gets a Boolean value that specifies whether the meta property is required. Read-only.
- `Type As MsoMetaPropertyType  (read-only)`  
  Gets the data type of a MetaProperty object. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the MetaProperty object. Read-only.

## Methods (1)

- `Validate() As String`  
  Validates a MetaProperty object representing a single property value according to a schema.
