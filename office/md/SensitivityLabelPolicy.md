# SensitivityLabelPolicy

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {88FF5F69-FACF-4667-8DC8-A85B8225DF15}  

Represents sensitivity label policy of the user's organization.

**Remarks:** Sensitivity label policy must be initialized before the start of the document otherwise the document will not be able to use the sensitivity label policy set by the organization.

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

Application.SensitivityLabelPolicy.CompleteInitialize myInitInfo
```

## Properties (2)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SensitivityLabelPolicy object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SensitivityLabelPolicy object was created. Read-only.

## Methods (3)

- `BeginInitialize() As String`  
  Begins the SensitivityLabelPolicy initialization sequence.
- `CompleteInitialize(SensitivityLabelInitInfo As SensitivityLabelInitInfo)`  
  Completes the SensitivityLabelPolicy initialization sequence.
    - `SensitivityLabelInitInfo As SensitivityLabelInitInfo` (required): The sensitivity label policy initialization object.
- `CreateSensitivityLabelInitInfo() As SensitivityLabelInitInfo`  
  Creates a new SensitivityLabelInitInfo object that can be passed to the CompleteInitialize method.
