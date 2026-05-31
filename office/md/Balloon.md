# Balloon

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0324-0000-0000-C000-000000000046}  

## Properties (15)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Checkboxes As Object  (read-only)`
- `Labels As Object  (read-only)`
- `BalloonType As MsoBalloonType  (read/write)`
- `Icon As MsoIconType  (read/write)`
- `Heading As String  (read/write)`
- `Text As String  (read/write)`
- `Mode As MsoModeType  (read/write)`
- `Animation As MsoAnimationType  (read/write)`
- `Button As MsoButtonSetType  (read/write)`
- `Callback As String  (read/write)`
- `Private As Long  (read/write)`
- `Name As String  (read-only)`

## Methods (3)

- `SetAvoidRectangle(Left As Long, Top As Long, Right As Long, Bottom As Long)`
- `Show() As MsoBalloonButtonType`
- `Close()`
