# LabelInfo

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {DFD3BED7-93EC-4BCE-866C-6BAB41D28621}  

Represents the label information data object.

**Remarks:** The LabelInfo object can be passed to SetLabel method of SensitivityLabel object.

**Example:**

```vba
Sub SetLabelInfo()

 Dim myLabelInfo As Office.LabelInfo
 Set myLabelInfo = Application.ActiveDocument.SensitivityLabel.CreateLabelInfo()
 With myLabelInfo
  .ActionId = "5cc46055-305d-4bc1-8f5f-5edf82231378"
  .AssignmentMethod = MsoAssignmentMethod.PRIVILEGED
  .ContentBits = 4
  .IsEnabled = True
  .Justification = "Some justification needed only if downgrading label."
  .LabelId = "9203368f-916c-4d59-8292-9f1c6a1e8f39"
  .LabelName = "MyLabelName"
  .SetDate = Now()
  .SiteId = "6c15903a-880e-4e17-818a-6cb4f7935615"
 End With

End Sub
```

## Properties (11)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `LabelId As String  (read/write)`
- `LabelName As String  (read/write)`
- `AssignmentMethod As MsoAssignmentMethod  (read/write)`
- `Justification As String  (read/write)`
- `IsEnabled As Boolean  (read/write)`
- `SetDate As String  (read/write)`
- `SiteId As String  (read/write)`
- `ActionId As String  (read/write)`
- `ContentBits As Long  (read/write)`
