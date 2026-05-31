# SensitivityLabelInitInfo

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {DE9CD4FF-754A-49DD-A0DC-B787DA2DB0A1}  

Represents the sensitivity label policy initialization data object.

**Remarks:** The SensitivityLabelInitInfo object can be passed to CompleteInitialization method of SensitivityLabelPolicy object.

**Example:**

```vba
Function GetSensitivityLabelsPolicyXml(policyVersion As String)
 Dim myOrgPolicyInXml as String

 ' Set myOrgPolicyInXml based on your organization's policy based on policyVersion as XML.
 GetSensitivityLabelsPolicyXml = myOrgPolicyInXml
End Function

Dim supportedPolicyVersion As String
Dim myInitInfo As Office.SensitivityLabelInitInfo

supportedPolicyVersion = Application.SensitivityLabelPolicy.BeginInitialize

Set myInitInfo = Application.SensitivityLabelPolicy.CreateSensitivityLabelInitInfo()
myInitInfo.UserId = "someone@example.com"
myInitInfo.SensitivityLabelsPolicyXml = GetSensitivityLabelsPolicyXml(supportedPolicyVersion)
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SensitivityLabelInitInfo object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SensitivityLabelInitInfo object was created. Read-only.
- `UserId As String  (read/write)`  
  Use the UserId property to set the string representing the UPN of the identity of an Office Account signed into Office.
- `SensitivityLabelsPolicyXml As String  (read/write)`  
  Use the SensitivityLabelsPolicyXml property to set the user's organization sensitivity label policy data as XML formatted string.
