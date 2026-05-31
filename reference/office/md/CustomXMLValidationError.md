# CustomXMLValidationError

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB0E-0000-0000-C000-000000000046}  

Represents a single validation error in a CustomXMLValidationErrors collection.

**Remarks:** Validation errors can either be triggered when validating an operation against the schema, such as when adding a node, or triggered by the user when some condition is not met. For example, if a start date is later than an end date.

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

## Properties (8)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLValidationError object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLValidationError object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLValidationError object. Read-only.
- `Name As String  (read-only)`  
  Gets the name of an error in a CustomXMLValidationError object. If no errors exist, the property returns Nothing. Read-only.
- `Node As CustomXMLNode  (read-only)`  
  Gets a node in a CustomXMLValidationError object, if any exist. If no nodes exist, the property returns Nothing. Read-only.
- `Text As String  (read-only)`  
  Gets the text in the CustomXMLValidationError object. Read-only.
- `Type As MsoCustomXMLValidationErrorType  (read-only)`  
  Gets the type of error generated from the CustomXMLValidationError object. Read-only.
- `ErrorCode As Long  (read-only)`  
  Gets a number representing a validation error in a CustomXMLValidationError object. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the CustomXMLValidationError object representing a data validation error.
