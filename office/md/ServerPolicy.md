# ServerPolicy

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0390-0000-0000-C000-000000000046}  

Represents a policy specified for a document type stored on a server running Microsoft Office SharePoint Server.

**Remarks:** The ServerPolicy object is composed of individual PolicyItem objects representing the individual policy definitions for the active document.

**Example:**

```vba
Sub ListPolicyItems()
Dim objSrvPolicy As ServerPolicy
Dim objPolicyItem As PolicyItem
Dim strPolicyItemList As String

Set objSrvPolicy = ActiveDocument.ServerPolicy

For Each objPolicyItem In objSrvPolicy
 strPolicyItemList = "Policy Item " & objPolicyItem.Name & " - " & _
 objPolicyItem.Description & vbCrLf
Next
MsgBox (strPolicyItemList)

End Sub
```

## Properties (10)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ServerPolicy object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ServerPolicy object was created. Read-only.
- `Item As PolicyItem  (read-only)`  
  Gets a PolicyItem object from the ServerPolicy collection. Read-only.
- `Id As String  (read-only)`  
  Gets the ID of a server policy. Read-only.
- `Name As String  (read-only)`  
  Gets the name of the ServerPolicy object. Read-only.
- `Description As String  (read-only)`  
  A description of what the server policy is and its purpose. Read-only.
- `Statement As String  (read-only)`  
  Gets the information specified in the policy statement. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the ServerPolicy collection. Read-only.
- `BlockPreview As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether you can preview items by using this policy. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the ServerPolicy object. Read-only.
