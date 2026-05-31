# Interaction

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (22)

- `AppActivate(Title As Variant, [Wait As Variant])`
- `Beep()`
- `CreateObject(Class As String, [ServerName As String]) As Variant`
- `DoEvents() As Integer`
- `GetObject([PathName As Variant], [Class As Variant]) As Variant`
- `InputBox(Prompt As Variant, [Title As Variant], [Default As Variant], [XPos As Variant], [YPos As Variant], [HelpFile As Variant], [Context As Variant]) As String`
- `MsgBox(Prompt As Variant, [Buttons As VbMsgBoxStyle], [Title As Variant], [HelpFile As Variant], [Context As Variant]) As VbMsgBoxResult`
- `SendKeys(String As String, [Wait As Variant])`
- `Shell(PathName As Variant, [WindowStyle As VbAppWinStyle]) As Double`
- `Partition(Number As Variant, Start As Variant, Stop As Variant, Interval As Variant) As Variant`
- `Choose(Index As Single, Choice As SAFEARRAY(Variant)) As Variant`
- `_B_var_Environ(Expression As Variant) As Variant`
- `_B_str_Environ(Expression As Variant) As String`
- `Switch(VarExpr As SAFEARRAY(Variant)) As Variant`
- `_B_var_Command() As Variant`
- `_B_str_Command() As String`
- `IIf(Expression As Variant, TruePart As Variant, FalsePart As Variant) As Variant`
- `GetSetting(AppName As String, Section As String, Key As String, [Default As Variant]) As String`
- `SaveSetting(AppName As String, Section As String, Key As String, Setting As String)`
- `DeleteSetting(AppName As String, [Section As Variant], [Key As Variant])`
- `GetAllSettings(AppName As String, Section As String) As Variant`
- `CallByName(Object As Object, ProcName As String, CallType As VbCallType, Args As SAFEARRAY(Variant), lcid As Long) As Variant`  
  Support IDispatch::Invoke
