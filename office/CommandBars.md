# CommandBars

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {55F88893-7708-11D1-ACEB-006008961DA5}  

## Properties (16)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `ActionControl As CommandBarControl  (read-only)`
- `ActiveMenuBar As CommandBar  (read-only)`
- `Count As Long  (read-only)`
- `DisplayTooltips As Boolean  (read/write)`
- `DisplayKeysInTooltips As Boolean  (read/write)`
- `Item As CommandBar  (read-only)`
- `LargeButtons As Boolean  (read/write)`
- `MenuAnimationStyle As MsoMenuAnimation  (read/write)`
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`
- `AdaptiveMenus As Boolean  (read/write)`
- `DisplayFonts As Boolean  (read/write)`
- `DisableCustomize As Boolean  (read/write)`
- `DisableAskAQuestionDropdown As Boolean  (read/write)`

## Methods (13)

- `Add([Name As Variant], [Position As Variant], [MenuBar As Variant], [Temporary As Variant]) As CommandBar`
- `FindControl([Type As Variant], [Id As Variant], [Tag As Variant], [Visible As Variant]) As CommandBarControl`
- `ReleaseFocus()`
- `FindControls([Type As Variant], [Id As Variant], [Tag As Variant], [Visible As Variant]) As CommandBarControls`
- `ExecuteMso(idMso As String)`
- `GetEnabledMso(idMso As String) As Boolean`
- `GetVisibleMso(idMso As String) As Boolean`
- `GetPressedMso(idMso As String) As Boolean`
- `GetLabelMso(idMso As String) As String`
- `GetScreentipMso(idMso As String) As String`
- `GetSupertipMso(idMso As String) As String`
- `GetImageMso(idMso As String, Width As Long, Height As Long) As IPictureDisp`
- `CommitRenderingTransaction(hwnd As Long)`

## Events (1)

- `OnUpdate()`
