# CustomXMLParts

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB0C-0000-0000-C000-000000000046}  

Represents a collection of CustomXMLPart objects.

**Remarks:** There are three default parts that are always created with a document. These are cover pages, doc properties, and app properties. The last two were in previous versions of Microsoft Word but are now provided in XML form in the CustomXMLParts object collection.

**Example:**

```vba
Sub AddPartToCollection()
    Dim myPart As CustomXMLPart

    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>")

End Sub
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLParts object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLParts object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLParts object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the CustomXMLParts collection. Read-only.
- `Item As CustomXMLPart  (read-only)`  
  Gets a CustomXMLPart object from the CustomXMLParts collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Add([XML As String], [SchemaCollection As Variant]) As CustomXMLPart`  
  Allows you to add a new CustomXMLPart to a file.
    - `XML As String` (optional): Contains the XML to add to the newly created CustomXMLPart.
    - `SchemaCollection As Variant` (optional): Represents the set of schemas to be used to validate this stream.
- `SelectByID(Id As String) As CustomXMLPart`  
  Selects a custom XML part matching a GUID.
    - `Id As String` (required): Contains the GUID for the custom XML part.
- `SelectByNamespace(NamespaceURI As String) As CustomXMLParts`  
  Selects the collection of custom XML parts whose namespace matches the search criteria.
    - `NamespaceURI As String` (required): Contains a namespace URI.

## Events (3)

- `PartAfterAdd(NewPart As CustomXMLPart)`  
  Occurs just after a CustomXMLPart object is added to the CustomXMLParts collection.
    - `NewPart As CustomXMLPart` (required): The part that was added.
- `PartBeforeDelete(OldPart As CustomXMLPart)`  
  Occurs just before a CustomXMLPart object is deleted from the CustomXMLParts collection.
    - `OldPart As CustomXMLPart` (required): The part that is about to be deleted.
- `PartAfterLoad(Part As CustomXMLPart)`  
  Occurs just after a CustomXMLPart object is loaded.
    - `Part As CustomXMLPart` (required): The part that was loaded.
