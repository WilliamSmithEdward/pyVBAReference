# DocumentProperties

**Type:** Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {2DF8D04D-5BFA-101B-BDE5-00AA0044DE52}  

A collection of DocumentProperty objects. Each DocumentProperty object represents a built-in or custom property of a container document.

**Remarks:** Use the Add method to create a new custom property and add it to the DocumentProperties collection. You cannot use the Add method to create a built-in document property. Use BuiltinDocumentProperties(index), where _index_ is the index number of the built-in document property, to return a single DocumentProperty object that represents a specific built-in document property. Use CustomDocumentProperties(index), where _index_ is the number of the custom document property, to return a DocumentProperty object that represents a specific custom document property.

## Properties (6)

- `Parent As Object  (read-only)`  
  Gets the Parent object for the DocumentProperties object. Read-only.
- `Item As HRESULT  (read-only)`  
  Gets a DocumentProperty object from the DocumentProperties collection. Read-only.
- `Count As HRESULT  (read-only)`  
  Gets a Long indicating the number of items in the DocumentProperties collection. Read-only.
- `_NewEnum As HRESULT  (read-only)`
- `Application As HRESULT  (read-only)`  
  Gets an Application object that represents the container application for the DocumentProperties object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As HRESULT  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the DocumentProperties object was created. Read-only.

## Methods (1)

- `Add(Name As String, LinkToContent As Boolean, [Type As Variant], [Value As Variant], [LinkSource As Variant], lcid As Long, ppIDocProp As DocumentProperty)`  
  Creates a new custom document property. You can add a new document property only to the custom DocumentProperties collection.
    - `Name As String` (required): The string of the Name of the property.
    - `LinkToContent As Boolean` (required): Specifies whether the LinkToContent property is linked to the contents of the container document. If this argument is True, the _LinkSource_ argument is required; if it's False, the _Value_ argument is required.
    - `Type As Variant` (optional): The data type of the Type property. Can be one of the following MsoDocProperties constants: msoPropertyTypeBoolean, msoPropertyTypeDate, msoPropertyTypeFloat, msoPropertyTypeNumber, or msoPropertyTypeString.
    - `Value As Variant` (optional): The data value of the Value property, if it's not linked to the contents of the container document. The value is converted to match the data type specified by the _Type_ argument, and if it can't be converted, an error occurs. If _LinkToContent_ is True, the argument is ignored, and the new document property is assigned a default value until the linked property values are updated by the container application (usually when the document is saved).
    - `LinkSource As Variant` (optional): Ignored if _LinkToContent_ is False. The source of the LinkSource property. The container application determines what types of source linking you can use. For example, DDE links use the "Server|Document!Item" syntax.
