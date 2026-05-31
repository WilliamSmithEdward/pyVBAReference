# PolicyItem

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0391-0000-0000-C000-000000000046}  

Represents an item within a ServerPolicy object that contains the settings for one policy.

**Remarks:** A policy item cannot exist outside the scope of a policy. Policy items are distinct conditions defined for a document stored on a server running Microsoft Office SharePoint Server.

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

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the PolicyItem object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PolicyItem object was created. Read-only.
- `Id As String  (read-only)`  
  Gets the ID of a policy item. PolicyItem objects are contained in ServerPolicy objects. Read-only.
- `Name As String  (read-only)`  
  Gets the name of the PolicyItem object. Read-only.
- `Description As String  (read-only)`  
  Gets a description of the current state of the policy item. Read-only.
- `Data As String  (read-only)`  
  Gets the information that is used to implement the policy item. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the PolicyItem object. Read-only.
