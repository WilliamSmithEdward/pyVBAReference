# CustomXMLValidationErrors

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB0F-0000-0000-C000-000000000046}  

Represents a collection of CustomXMLValidationError objects.

**Example:**

```vba
Dim ValErrors As CustomXMLValidationErrors
Dim ValError As CustomXMLValidationError
Dim cxp1 As CustomXMLPart
Dim intError As Integer

On Error Go To validation_error

 With ActiveDocument

    ' Add and populate a custom xml part
    set cxp1 = .CustomXMLParts.Add "<invoice>"

    ' Add a node
    cxp1.AddNode "<quantity>", "supplier", "urn:invoice:namespace"

 End With

If ValErrors.Count > 0 then
   For Each ValError In ValErrors
      DeBug.Print("Error name: " & ValError.Name & " Error description: " & ValError.Text)
   Next
End If

Exit Sub

validation_error:
   CustomXMLValidationErrors.Add(ValError.Name, ValError.Text))
Resume
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLValidationErrors object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLValidationErrors object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLValidationErrors object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the CustomXMLValidationErrors collection. Read-only.
- `Item As CustomXMLValidationError  (read-only)`  
  Gets a CustomXMLValidationError object from the CustomXMLValidationErrors collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Node As CustomXMLNode, ErrorName As String, [ErrorText As String], [ClearedOnUpdate As Boolean])`  
  Adds a CustomXMLValidationError object containing an XML validation error to the CustomXMLValidationErrors collection.
    - `Node As CustomXMLNode` (required): Represents the node where the error occurred.
    - `ErrorName As String` (required): Contains the name of the error.
    - `ErrorText As String` (optional): Contains the descriptive error text.
    - `ClearedOnUpdate As Boolean` (optional): Specifies whether the error is to be cleared from the CustomXMLValidationErrors collection when the XML is corrected and updated.
