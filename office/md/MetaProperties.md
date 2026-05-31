# MetaProperties

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C038E-0000-0000-C000-000000000046}  

Represents a collection of properties describing the metadata stored in a document.

**Example:**

```vba
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String
Dim result As String

result = metaProps(1).Validate
ValidateMetaProperty = result
End Function
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the MetaProperties object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the MetaProperties object was created. Read-only.
- `Item As MetaProperty  (read-only)`  
  Gets a MetaProperty object from the MetaProperties collection. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the MetaProperties collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the MetaProperties object. Read-only.
- `SchemaXml As String  (read-only)`  
  Gets the schema XML for the MetaProperties object. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `GetItemByInternalName(InternalName As String) As MetaProperty`  
  Gets a property's value specifying its name as opposed to its index value.
    - `InternalName As String` (required): Contains the name of the property.
- `Validate() As String`  
  Validates all of the properties in a MetaProperties collection object according to a schema.
