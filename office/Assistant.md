# Assistant

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0322-0000-0000-C000-000000000046}  

## Properties (26)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Top As Long  (read/write)`
- `Left As Long  (read/write)`
- `NewBalloon As Balloon  (read-only)`
- `BalloonError As MsoBalloonErrorType  (read-only)`
- `Visible As Boolean  (read/write)`
- `Animation As MsoAnimationType  (read/write)`
- `Reduced As Boolean  (read/write)`
- `AssistWithHelp As Boolean  (read/write)`
- `AssistWithWizards As Boolean  (read/write)`
- `AssistWithAlerts As Boolean  (read/write)`
- `MoveWhenInTheWay As Boolean  (read/write)`
- `Sounds As Boolean  (read/write)`
- `FeatureTips As Boolean  (read/write)`
- `MouseTips As Boolean  (read/write)`
- `KeyboardShortcutTips As Boolean  (read/write)`
- `HighPriorityTips As Boolean  (read/write)`
- `TipOfDay As Boolean  (read/write)`
- `GuessHelp As Boolean  (read/write)`
- `SearchWhenProgramming As Boolean  (read/write)`
- `Item As String  (read-only)`
- `FileName As String  (read/write)`
- `Name As String  (read-only)`
- `On As Boolean  (read/write)`

## Methods (7)

- `Move(xLeft As Long, yTop As Long)`
- `Help()`
- `StartWizard(On As Boolean, Callback As String, PrivateX As Long, [Animation As Variant], [CustomTeaser As Variant], [Top As Variant], [Left As Variant], [Bottom As Variant], [Right As Variant]) As Long`
- `EndWizard(WizardID As Long, varfSuccess As Boolean, [Animation As Variant])`
- `ActivateWizard(WizardID As Long, act As MsoWizardActType, [Animation As Variant])`
- `ResetTips()`
- `DoAlert(bstrAlertTitle As String, bstrAlertText As String, alb As MsoAlertButtonType, alc As MsoAlertIconType, ald As MsoAlertDefaultType, alq As MsoAlertCancelType, varfSysAlert As Boolean) As Long`
