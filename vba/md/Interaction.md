# Interaction

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (22)

- `AppActivate(Title As Variant, [Wait As Variant])`  
  Activates an application window.
    - `Title As Variant` (required): Required. String expression specifying the title in the title bar of the application window you want to activate. The task ID returned by the Shell function can be used in place of _title_ to activate an application.
    - `Wait As Variant` (optional): Optional. Boolean value specifying whether the calling application has the focus before activating another. If False (default), the specified application is immediately activated, even if the calling application does not have the focus. If True, the calling application waits until it has the focus, and then activates the specified application.
- `Beep()`  
  Sounds a tone through the computer's speaker.
- `CreateObject(Class As String, [ServerName As String]) As Variant`  
  Creates and returns a reference to an ActiveX object.
    - `Class As String` (required): Required; Variant (String). The application name and class of the object to create.
    - `ServerName As String` (optional): Optional; Variant (String). The name of the network server where the object will be created. If _servername_ is an empty string (""), the local machine is used.
- `DoEvents() As Integer`  
  Yields execution so that the operating system can process other events.
- `GetObject([PathName As Variant], [Class As Variant]) As Variant`  
  Returns a reference to an object provided by an ActiveX component.
    - `PathName As Variant` (optional): Optional; Variant (String). The full path and name of the file containing the object to retrieve. If _pathname_ is omitted, _class_ is required.
    - `Class As Variant` (optional): Optional; Variant (String). A string representing the class of the object.
- `InputBox(Prompt As Variant, [Title As Variant], [Default As Variant], [XPos As Variant], [YPos As Variant], [HelpFile As Variant], [Context As Variant]) As String`  
  Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns a String containing the contents of the text box.
    - `Prompt As Variant` (required): Required. String expression displayed as the message in the dialog box. The maximum length of _prompt_ is approximately 1024 characters, depending on the width of the characters used. If _prompt_ consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return-linefeed character combination ((Chr(13) & (Chr(10)) between each line.
    - `Title As Variant` (optional): Optional. String expression displayed in the title bar of the dialog box. If you omit _title_, the application name is placed in the title bar.
    - `Default As Variant` (optional): Optional. String expression displayed in the text box as the default response if no other input is provided. If you omit _default_, the text box is displayed empty.
    - `XPos As Variant` (optional): Optional. Numeric expression that specifies, in twips, the horizontal distance of the left edge of the dialog box from the left edge of the screen. If _xpos_ is omitted, the dialog box is horizontally centered.
    - `YPos As Variant` (optional): Optional. Numeric expression that specifies, in twips, the vertical distance of the upper edge of the dialog box from the top of the screen. If _ypos_ is omitted, the dialog box is vertically positioned approximately one-third of the way down the screen.
    - `HelpFile As Variant` (optional): Optional. String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If _helpfile_ is provided, _context_ must also be provided.
    - `Context As Variant` (optional): Optional. Numeric expression that is the Help context number assigned to the appropriate Help topic by the Help author. If _context_ is provided, _helpfile_ must also be provided.
- `MsgBox(Prompt As Variant, [Buttons As VbMsgBoxStyle], [Title As Variant], [HelpFile As Variant], [Context As Variant]) As VbMsgBoxResult`  
  Displays a message in a dialog box, waits for the user to click a button, and returns an Integer indicating which button the user clicked.
    - `Prompt As Variant` (required): Required. String expression displayed as the message in the dialog box. The maximum length of _prompt_ is approximately 1024 characters, depending on the width of the characters used. If _prompt_ consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line.
    - `Buttons As VbMsgBoxStyle` (optional): Optional. Numeric expression that is the combination of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for _buttons_ is 0.
    - `Title As Variant` (optional): Optional. String expression displayed in the title bar of the dialog box. If you omit _title_, the application name is placed in the title bar.
    - `HelpFile As Variant` (optional): Optional. String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If _helpfile_ is provided, _context_ must also be provided.
    - `Context As Variant` (optional): Optional. Numeric expression that is the Help context number assigned to the appropriate Help topic by the Help author. If _context_ is provided, _helpfile_ must also be provided.
- `SendKeys(String As String, [Wait As Variant])`  
  Sends one or more keystrokes to the active window as if typed at the keyboard.
    - `String As String` (required): Required. String expression specifying the keystrokes to send.
    - `Wait As Variant` (optional): Optional. Boolean value specifying the wait mode. If False (default), control is returned to the procedure immediately after the keys are sent. If True, keystrokes must be processed before control is returned to the procedure.
- `Shell(PathName As Variant, [WindowStyle As VbAppWinStyle]) As Double`  
  Runs an executable program and returns a Variant (Double) representing the program's task ID if successful; otherwise, it returns zero.
    - `PathName As Variant` (required): Required; Variant (String). Name of the program to execute and any required arguments or command-line switches; may include directory or folder and drive. On the Macintosh, you can use the MacID function to specify an application's signature instead of its name. The following example uses the signature for Microsoft Word: Shell MacID("MSWD")
    - `WindowStyle As VbAppWinStyle` (optional): Optional. Variant (Integer) corresponding to the style of the window in which the program is to be run. If _windowstyle_ is omitted, the program is started minimized with focus. On the Macintosh (System 7.0 or later), _windowstyle_ only determines whether or not the application gets the focus when it is run.
- `Partition(Number As Variant, Start As Variant, Stop As Variant, Interval As Variant) As Variant`  
  Returns a Variant (String) indicating where a number occurs within a calculated series of ranges.
    - `Number As Variant` (required): Required. The number that you want to evaluate against the ranges.
    - `Start As Variant` (required): Required. The number that is the start of the overall range of numbers. The number can't be less than 0.
    - `Stop As Variant` (required): Required. The number that is the end of the overall range of numbers. The number can't be equal to or less than _start_.
    - `Interval As Variant` (required): Required. The number that is the difference between one range and the next. The number can't be less than 1.
- `Choose(Index As Single, Choice As SAFEARRAY(Variant)) As Variant`  
  Selects and returns a value from a list of arguments.
    - `Index As Single` (required): Required. Numeric expression or field that results in a value between 1 and the number of available choices.
    - `Choice As SAFEARRAY(Variant)` (required): Required. Variant expression containing one of the possible choices.
- `_B_var_Environ(Expression As Variant) As Variant`
- `_B_str_Environ(Expression As Variant) As String`
- `Switch(VarExpr As SAFEARRAY(Variant)) As Variant`  
  Evaluates a list of expressions and returns a Variant value or an expression associated with the first expression in the list that is True.
- `_B_var_Command() As Variant`
- `_B_str_Command() As String`
- `IIf(Expression As Variant, TruePart As Variant, FalsePart As Variant) As Variant`  
  Returns one of two parts, depending on the evaluation of an expression.
    - `TruePart As Variant` (required): Required. Value or expression returned if _expr_ is True.
    - `FalsePart As Variant` (required): Required. Value or expression returned if _expr_ is False.
- `GetSetting(AppName As String, Section As String, Key As String, [Default As Variant]) As String`  
  Returns a key setting value from an application's entry in the Windows registry or (on the Macintosh) information in the application's initialization file.
    - `AppName As String` (required): Required. String expression containing the name of the application or project whose key setting is requested. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.
    - `Section As String` (required): Required. String expression containing the name of the section where the key setting is found.
    - `Key As String` (required): Required. String expression containing the name of the key setting to return.
    - `Default As Variant` (optional): Optional. Expression containing the value to return if no value is set in the key setting. If omitted, _default_ is assumed to be a zero-length string ("").
- `SaveSetting(AppName As String, Section As String, Key As String, Setting As String)`  
  Saves or creates an application entry in the application's entry in the Windows registry or (on the Macintosh) information in the application's initialization file.
    - `AppName As String` (required): Required. String expression containing the name of the application or project to which the setting applies. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.
    - `Section As String` (required): Required. String expression containing the name of the section where the key setting is being saved.
    - `Key As String` (required): Required. String expression containing the name of the key setting being saved.
    - `Setting As String` (required): Required. Expression containing the value that _key_ is being set to.
- `DeleteSetting(AppName As String, [Section As Variant], [Key As Variant])`  
  Deletes a section or key setting from an application's entry in the Windows registry or (on the Macintosh) information in the application's initialization file.
    - `AppName As String` (required): Required. String expression containing the name of the application or project to which the section or key setting applies. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.
    - `Section As Variant` (optional): Required. String expression containing the name of the section where the key setting is being deleted. If only _appname_ and _section_ are provided, the specified section is deleted along with all related key settings.
    - `Key As Variant` (optional): Optional. String expression containing the name of the key setting being deleted.
- `GetAllSettings(AppName As String, Section As String) As Variant`  
  Returns a list of key settings and their respective values (originally created with SaveSetting) from an application's entry in the Windows registry or (on the Macintosh) information in the application's initialization file.
    - `AppName As String` (required): Required. String expression containing the name of the application or project whose key settings are requested. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.
    - `Section As String` (required): Required. String expression containing the name of the section whose key settings are requested. GetAllSettings returns a Variant whose contents is a two-dimensional array of strings containing all the key settings in the specified section and their corresponding values.
- `CallByName(Object As Object, ProcName As String, CallType As VbCallType, Args As SAFEARRAY(Variant), lcid As Long) As Variant`  
  Executes a method of an object, or sets or returns a property of an object.
    - `Object As Object` (required): Required: Variant (Object). The name of the object on which the function will be executed.
    - `ProcName As String` (required): Required: Variant (String). A string expression containing the name of a property or method of the object.
    - `CallType As VbCallType` (required): Required: Constant. A constant of type vbCallType representing the type of procedure being called.
