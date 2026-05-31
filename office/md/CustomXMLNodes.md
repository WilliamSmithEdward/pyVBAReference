# CustomXMLNodes

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB03-0000-0000-C000-000000000046}  

Contains a collection of CustomXMLNode objects representing the XML nodes in a document.

**Remarks:** The Attributes and the ChildNodes properties return collections of nodes of this type.

**Example:**

```vba
Sub CustomXmlNodes()
    Dim cxp1 As CustomXMLPart
    Dim cxns As CustomXMLNodes

    With ActiveDocument

        ' Returns the first custom xml part with the given root namespace.
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")

        ' Get custom xml nodes using XPath.
        Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")

    End With

End Sub
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLNodes object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLNodes object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLNodes object. Read-only.
- `Count As Long  (read-only)`  
  Gets a count of the number of CustomXMLNode objects in a CustomXMLNodes collection. Read-only.
- `Item As CustomXMLNode  (read-only)`  
  Gets a CustomXMLNode object from the CustomXMLNodes collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
