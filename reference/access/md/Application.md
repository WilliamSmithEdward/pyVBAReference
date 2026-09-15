# Application

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}  

The Application object refers to the active Microsoft Access application.

**Remarks:** The Application object contains all Access objects and collections. Use the Application object to apply methods or property settings to the entire Access application. For example, you can use the SetOption method of the Application object to set database options from Visual Basic. The following example shows how you can set the Display Status Bar check box on the Current Database tab of the Access Options dialog box. Access is a COM component that supports Automation, formerly called OLE Automation. You can manipulate Access objects from another application that also supports Automation. To do this, you use the Application object. For example, Microsoft Visual Basic is a COM component. You can open an Access database from Visual Basic and work with its objects.

## Properties (42)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `CodeContextObject As Object  (read-only)`  
  Use the CodeContextObject property to determine the object in which a macro or Visual Basic code is executing. Read-only Object.
- `MenuBar As String  (read/write)`  
  Specifies a custom menu to display for a Microsoft Access database. Read/write String.
- `CurrentObjectType As AcObjectType  (read-only)`  
  Use the CurrentObjectType property together with the Application object to determine the type of the active database object (table, query, form, report, macro, module, server view, database diagram, or stored procedure). The active database object is the object that has the focus or in which code is running. Read-only AcObjectType.
- `CurrentObjectName As String  (read-only)`  
  Use the CurrentObjectName property with the Application object to determine the name of the active database object. The active database object is the object that has the focus or in which code is running. Read-only String.
- `Forms As Forms  (read-only)`  
  Use the Forms property to return a read-only reference to the Forms collection and its related properties.
- `Reports As Reports  (read-only)`  
  Use the Reports property to access the read-only Reports collection and its related properties.
- `Screen As Screen  (read-only)`  
  Use the Screen property to return a reference the Screen object and its related properties. Read-only.
- `DoCmd As DoCmd  (read-only)`  
  Use the DoCmd property to access the read-only DoCmd object and its related methods. Read-only DoCmd.
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `UserControl As Boolean  (read/write)`  
  Use the UserControl property to determine whether the current Microsoft Access application was started by the user or by another application with Automation, formerly called OLE Automation. Read/write Boolean.
- `DBEngine As DBEngine  (read-only)`  
  Use the DBEngine property in Visual Basic to access the current DBEngine object and its related properties. Read-only DBEngine.
- `CommandBars As CommandBars  (read-only)`  
  Use the CommandBars property to return a reference to the CommandBars collection object. Read-only CommandBars object.
- `References As References  (read-only)`  
  Use the References property to access the References collection and its related properties, methods, and events. Read-only References collection.
- `Modules As Modules  (read-only)`  
  Use the Modules property to access the Modules collection and its related properties. Read-only Modules object.
- `IsCompiled As Boolean  (read-only)`  
  The IsCompiled property returns a Boolean value indicating whether the Visual Basic project is in a compiled state. Read-only Boolean.
- `VBE As VBE  (read-only)`  
  Use the VBE property to return a reference to the current VBE object and its related properties. The VBE property of the Application object represents the Microsoft Visual Basic for Applications editor. Read-only VBE object.
- `CurrentProject As CurrentProject  (read-only)`  
  Use the CurrentProject property to access the CurrentProject object and its related collections, properties, and methods. Read-only CurrentProject object.
- `CurrentData As CurrentData  (read-only)`  
  Use the CurrentData property to access the CurrentData object and its related collections. Read-only CurrentData object.
- `CodeProject As CodeProject  (read-only)`  
  Use the CodeProject property to access the CodeProject object and its related collections, properties, and methods. Read-only CodeProject object.
- `CodeData As CodeData  (read-only)`  
  Use the CodeData property to access the CodeData object and its related collections. Read-only CodeData object.
- `ProductCode As String  (read-only)`  
  Use the ProductCode property to determine the Microsoft Access globally unique identifier (GUID). Read-only String.
- `COMAddIns As COMAddIns  (read-only)`  
  Use the COMAddIns property to return a reference to the current COMAddIns collection object and its related properties. Read-only COMAddIns object.
- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.
- `LanguageSettings As LanguageSettings  (read-only)`  
  Use the LanguageSettings property to return a read-only reference to the current LanguageSettings object and its related properties.
- `FeatureInstall As MsoFeatureInstall  (read/write)`  
  Use the FeatureInstall property to specify or determine how Microsoft Access handles calls to methods and properties that require features not yet installed. Read/write MsoFeatureInstall.
- `FileDialog As FileDialog  (read-only)`  
  Returns a FileDialog object that represents a single instance of a file dialog box. Read-only.
- `BrokenReference As Boolean  (read-only)`  
  Returns a Boolean indicating whether the current database has any broken references to databases or type libraries. True if there are any broken references. Read-only.
- `Printers As Printers  (read-only)`  
  Returns the Printers collection representing all the available printers on the current system. Read-only Printers collection.
- `Printer As _Printer  (read/write)`  
  Returns or sets a Printer object representing the default printer on the current system. Read/write.
- `Version As String  (read-only)`  
  Returns a String indicating the version number of the currently installed copy of Microsoft Access. Read-only.
- `Build As Long  (read-only)`  
  Returns as a Long representing the build number of the currently installed copy of Microsoft Access. Read-only.
- `NewFileTaskPane As NewFile  (read-only)`  
  Returns a NewFile object that represents a document listed on the New File task pane. Read-only NewFile object.
- `AutoCorrect As _AutoCorrect  (read-only)`  
  Returns an AutoCorrect object that represents the AutoCorrect settings for Microsoft Access. Read-only.
- `AutomationSecurity As MsoAutomationSecurity  (read/write)`  
  Returns or sets an MsoAutomationSecurity constant that represents the security mode that Microsoft Access uses when it is programmatically opening files. Read/write.
- `MacroError As MacroError  (read-only)`  
  Returns a MacroError object that contains information about the latest error to occur in a macro. Read-only.
- `TempVars As TempVars  (read-only)`  
  Returns the collection of TempVar objects. Read-only TempVars.
- `Assistance As IAssistance  (read-only)`  
  Returns an IAssistance object that represents the Microsoft Office Help Viewer. Read-only.
- `WebServices As WebServices  (read-only)`  
  Gets the collection of installed Data Service data connections. Read-only WebServices.
- `ReturnVars As ReturnVars  (read-only)`  
  Returns the ReturnVars collection representing all the available ReturnVar variables. Read-only ReturnVars collection.

## Methods (79)

- `GetOption(OptionName As String) As Variant`  
  The GetOption method returns the current value of an option in the Access Options dialog box, available by choosing the Microsoft Office button, and then choosing Access Options. Variant.
    - `OptionName As String` (required): The name of the option. For a list of option name argument strings, see Set options from Visual Basic.
- `SetOption(OptionName As String, Setting As Variant)`  
  The SetOption method sets the current value of an option in the Access Options dialog box.
    - `OptionName As String` (required): The name of the option. For a list of option name argument strings, see Set options from Visual Basic.
    - `Setting As Variant` (required): A value corresponding to the option setting. The value of the setting argument depends on the possible settings for a particular option.
- `Echo(EchoOn As Integer, [bstrStatusBarText As String])`  
  The Echo method specifies whether Microsoft Access repaints the display screen.
    - `EchoOn As Integer` (required): True (default) indicates that the screen is repainted.
    - `bstrStatusBarText As String` (optional): A string expression that specifies the text to display in the status bar when repainting is turned on or off.
- `CloseCurrentDatabase()`  
  Use the CloseCurrentDatabase method to close the current database, either a Microsoft Access database or an Access project (.adp) from another application that has opened a database through Automation.
- `Quit([Option As AcQuitOption])`  
  The Quit method quits Microsoft Access. You can select one of several options for saving a database object before quitting.
- `SysCmd(Action As AcSysCmdAction, [Argument2 As Variant], [Argument3 As Variant]) As Variant`  
  Use the SysCmd method to display a progress meter or optional specified text in the status bar, return information about Microsoft Access and its associated files, or return the state of a specified database object (to indicate whether the object is open, is a new object, or has been changed but not saved). Variant.
    - `Action As AcSysCmdAction` (required): An AcSysCmdAction constant that identifies the type of action to take. This set of constants applies to a progress meter. The SysCmd method returns a Null if these actions are successful. Otherwise, Access generates a run-time error.
    - `Argument2 As Variant` (optional): The text to be displayed left-aligned in the status bar. This argument is required when the _Action_ argument is acSysCmdInitMeter, acSysCmdUpdateMeter, or acSysCmdSetStatus; this argument isn't valid for other _Action_ argument values. NOTE: When you specify the acSysCmdGetObjectState value for the _Action_ parameter, you must specify the appropriate AcObjectType constant.
    - `Argument3 As Variant` (optional): A numeric expression that controls the display of the progress meter. This argument is required when the _Action_ argument is acSysCmdInitMeter; this argument isn't valid for other _Action_ argument values. NOTE: When you specify the acSysCmdGetObjectState value for the _Action_ parameter, you must specify the name of the database object.
- `CreateForm([Database As Variant], [FormTemplate As Variant]) As Form`  
  The CreateForm method creates a form and returns a Form object.
    - `Database As Variant` (optional): The name of the database that contains the form template that you want to use to create a form. If you want the current database, omit this argument. If you want to use an open library database, specify the library database with this argument.
    - `FormTemplate As Variant` (optional): The name of the form that you want to use as a template to create a new form.
- `CreateReport([Database As Variant], [ReportTemplate As Variant]) As Report`  
  The CreateReport method creates a report and returns a Report object. For example, suppose you are building a custom wizard to create a sales report. Use the CreateReport method in your wizard to create a new report based on a specified report template.
    - `Database As Variant` (optional): The name of the database that contains the report template that you want to use to create a report. If you want the current database, omit this argument. If you want to use an open library database, specify the library database with this argument.
    - `ReportTemplate As Variant` (optional): The name of the report that you want to use as a template to create a new report.
- `DeleteControl(FormName As String, ControlName As String)`  
  The DeleteControl method deletes a specified control from a form.
    - `FormName As String` (required): The name of the form containing the control that you want to delete.
    - `ControlName As String` (required): The name of the control that you want to delete.
- `DeleteReportControl(ReportName As String, ControlName As String)`  
  The DeleteReportControl method deletes a specified control from a report.
    - `ReportName As String` (required): The name of the report containing the control that you want to delete.
    - `ControlName As String` (required): The name of the control that you want to delete.
- `CreateGroupLevel(ReportName As String, Expression As String, Header As Integer, Footer As Integer) As Long`  
  Use the CreateGroupLevel method to specify a field or expression on which to group or sort data in a report.
    - `ReportName As String` (required): The name of the report that will contain the new group level.
    - `Expression As String` (required): The field or expression to sort or group on.
    - `Header As Integer` (required): Indicates that a field or expression will have an associated group header. If the _Header_ argument is True (1), the field or expression will have a group header. If the _Header_ argument is False (0), the field or expression won't. You can create a header by setting the argument to True.
    - `Footer As Integer` (required): Indicates a field or expression will have an associated group footer. If the _Footer_ argument is True (1), the field or expression will have a group footer. If the _Footer_ argument is False (0), the field or expression won't. You can create a footer by setting the argument to True.
- `DMin(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DMin function to determine the minimum value in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the field for which you want to find the minimum or maximum value. It can be a string expression identifying a field in a table or query, or it can be an expression that performs calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DMin function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DMin function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DMin function returns a Null.
- `DMax(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DMax function to determine the maximum value in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the field for which you want to find the minimum or maximum value. It can be a string expression identifying a field in a table or query, or it can be an expression that performs calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DMax function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DMax function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DMax function returns a Null.
- `DSum(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DSum function to calculate the sum of a set of values in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the numeric field whose values you want to total. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DSum function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DSum function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DSum function returns a Null.
- `DAvg(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DAvg function to calculate the average of a set of values in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the field containing the numeric data that you want to average. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a calculation on data in that field. In _Expr__, you can include the name of a field in a table, a control on a form, a constant, or a function. If _Expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DAvg function is performed. For example, _Criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _Criteria_ is omitted, the DAvg function evaluates _Expr_ against the entire domain. Any field that is included in _Criteria_ must also be a field in _Domain_; otherwise, the DAvg function returns a Null.
- `DLookup(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DLookup function to get the value of a particular field from a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the field whose value you want to return. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DLookup function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DLookup function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DLookup function returns a Null.
- `DLast(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DLast function to return a random record from a particular field in a table or query when you need any value from that field.
    - `Expr As String` (required): An expression that identifies the field from which you want to find the first or last value. It can be either a string expression identifying a field in a table or query, or an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DLast function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the wrd WHERE. If _criteria_ is omitted, the DLast function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DLast function returns a Null.
- `DVar(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Estimates the variance across a sample in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the numeric field on which you want to find the variance. It can be a string expression identifying a field from a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function. Any field included in _expr_ must be a numeric field.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DVar function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DVar function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DVar function returns a Null.
- `DVarP(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Calculates the variance of a population in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the numeric field on which you want to find the variance. It can be a string expression identifying a field from a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function. Any field included in _expr_ must be a numeric field.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DVarP function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DVarP function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DVarP function returns a Null.
- `DStDev(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Estimates the standard deviation across a population sample in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the numeric field on which you want to find the standard deviation. It can be a string expression identifying a field from a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DStDev function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DStDev function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DStDev function will return a Null.
- `DStDevP(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Estimates the standard deviation across a population in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the numeric field on which you want to find the standard deviation. It can be a string expression identifying a field from a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DStDevP function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DStDevP function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DStDevP function will return a Null.
- `DFirst(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DFirst function to return a random record from a particular field in a table or query when you need any value from that field.
    - `Expr As String` (required): An expression that identifies the field from which you want to find the first or last value. It can be either a string expression identifying a field in a table or query, or an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression identifying the set of records that constitutes the domain.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DFirst function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the wrd WHERE. If _criteria_ is omitted, the DFirst function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DFirst function returns a Null.
- `DCount(Expr As String, Domain As String, [Criteria As Variant]) As Variant`  
  Use the DCount function to determine the number of records that are in a specified set of records (a domain).
    - `Expr As String` (required): An expression that identifies the field for which you want to count records. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a calculation on data in that field. In _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.
    - `Domain As String` (required): A string expression that identifies the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.
    - `Criteria As Variant` (optional): An optional string expression used to restrict the range of data on which the DCount function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the DCount function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the DCount function returns a Null.
- `Eval(StringExpr As String) As Variant`  
  Use the Eval function to evaluate an expression that results in a text string or a numeric value.
    - `StringExpr As String` (required): An expression that evaluates to an alphanumeric text string. For example, _stringexpr_ can be a function that returns a string or a numeric value, or it can be a reference to a control on a form. The _stringexpr_ argument must evaluate to a string or numeric value; it can't evaluate to a Microsoft Access object.
- `CurrentUser() As String`  
  Use the CurrentUser method to return the name of the current user of the database.
- `DDEInitiate(Application As String, Topic As String) As Variant`  
  Use the DDEInitiate function to begin a dynamic data exchange (DDE) conversation with another application. The DDEInitiate function opens a DDE channel for transfer of data between a DDE server and a client application.
    - `Application As String` (required): A string expression identifying an application that can participate in a DDE conversation. Usually, the _application_ argument is the name of an .exe file (without the .exe extension) for a Microsoft Windows-based application, such as Microsoft Excel.
    - `Topic As String` (required): A string expression that is the name of a topic recognized by the _application_ argument. Check the application's documentation for a list of topics.
- `DDEExecute(ChanNum As Variant, Command As String)`  
  Use the DDEExecute statement to send a command from a client application to a server application over an open dynamic data exchange (DDE) channel.
    - `ChanNum As Variant` (required): A channel number, the long integer returned by the DDEInitiate function.
    - `Command As String` (required): A command recognized by the server application. Check the server application's documentation for a list of these commands.
- `DDEPoke(ChanNum As Variant, Item As String, Data As String)`  
  Use the DDEPoke statement to supply text data from a client application to a server application over an open dynamic data exchange (DDE) channel.
    - `ChanNum As Variant` (required): A channel number, an integer returned by the DDEInitiate function.
    - `Item As String` (required): The name of a data item recognized by the application specified by the DDEInitiate function. Check the application's documentation for a list of possible items.
    - `Data As String` (required): The data to be supplied to the other application.
- `DDERequest(ChanNum As Variant, Item As String) As String`  
  Use the DDERequest function over an open dynamic data exchange (DDE) channel to request an item of information from a DDE server application.
    - `ChanNum As Variant` (required): A channel number, the integer returned by the DDEInitiate function.
    - `Item As String` (required): A string expression that's the name of a data item recognized by the application specified by the DDEInitiate function. Check the application's documentation for a list of possible items.
- `DDETerminate(ChanNum As Variant)`  
  Use the DDETerminate statement to close a specified dynamic data exchange (DDE) channel.
    - `ChanNum As Variant` (required): A channel number to close; refers to a channel opened by the DDEInitiate function.
- `DDETerminateAll()`  
  Use the DDETerminateAll statement to close all open dynamic data exchange (DDE) channels.
- `CurrentDb() As Database`  
  The CurrentDb method returns an object variable of type Database that represents the database currently open in the Microsoft Access window.
- `CodeDb() As Database`  
  Use the CodeDb method in a code module to determine the name of the Database object that refers to the database in which code is currently running. Use the CodeDb method to access Data Access Objects (DAO) that are part of a library database.
- `BuildCriteria(Field As String, FieldType As Integer, Expression As String) As String`  
  The BuildCriteria method returns a parsed criteria string as it would appear in the query design grid, in Filter By Form or Server Filter By Form mode. For example, you may want to set a form's Filter or ServerFilter property based on varying criteria from the user. Use the BuildCriteria method to construct the string expression argument for the Filter or ServerFilter property. String.
    - `Field As String` (required): The field for which you wish to define criteria.
    - `FieldType As Integer` (required): An intrinsic constant denoting the data type of the field. Can be set to one of the DAO DataTypeEnum values.
    - `Expression As String` (required): A string expression identifying the criteria to be parsed.
- `DefaultWorkspaceClone() As Workspace`  
  Use the DefaultWorkspaceClone method to create a new Workspace object without requiring the user to sign in again. For example, if you need to conduct two sets of transactions simultaneously in separate workspaces, you can use the DefaultWorkspaceClone method to create a second Workspace object with the same user name and password without prompting the user for this information again.
- `RefreshTitleBar()`  
  The RefreshTitleBar method refreshes the Microsoft Access title bar after the AppTitle or AppIcon property has been set in Visual Basic.
- `hWndAccessApp() As Long`  
  Use the hWndAccessApp method to determine the handle assigned by Windows to the main Microsoft Access window.
- `Run(Procedure As String, [Arg1 As Variant], [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Variant`  
  Use the Run method to carry out a specified Microsoft Access or user-defined Function or Sub procedure. Variant.
    - `Procedure As String` (required): The name of the Function or Sub procedure to be run. If you are calling a procedure in another database, use the project name and the procedure name separated by a dot in the form: "_projectname_._procedurename_" If you execute Visual Basic code containing the Run method in a library database, Access looks for the procedure first in the library database, and then in the current database.
- `Nz(Value As Variant, [ValueIfNull As Variant]) As Variant`  
  Use the Nz function to return zero (0), a zero-length string (""), or another specified value when a Variant is Null. For example, you can use this function to convert a Null value to another value and prevent it from propagating through an expression.
    - `Value As Variant` (required): A variable of data type Variant.
    - `ValueIfNull As Variant` (optional): Optional. A Variant that supplies a value to be returned if the variant argument is Null. This argument enables you to return a value other than zero or a zero-length string. NOTE: If you use the Nz function in an expression in a query without using the _ValueIfNull_ argument, the results will be a zero-length string in the fields that contain Null values.
- `LoadPicture(FileName As String) As Object`  
  The LoadPicture method loads a graphic into an ActiveX control.
    - `FileName As String` (required): The file name of the graphic to be loaded. The graphic can be a bitmap file (.bmp), icon file (.ico), run-length encoded file (.rle), or metafile (.wmf).
- `AccessError(ErrorNumber As Variant) As Variant`  
  Use the AccessError method to return the descriptive string associated with a Microsoft Access or Data Access Objects (DAO) error.
    - `ErrorNumber As Variant` (required): The number of the error for which you wish to return a descriptive string.
- `StringFromGUID(Guid As Variant) As Variant`  
  The StringFromGUID function converts a GUID, which is an array of type Byte, to a string.
    - `Guid As Variant` (required): An array of Byte data used to uniquely identify an application, component, or item of data to the operating system.
- `GUIDFromString(String As Variant) As Variant`  
  The GUIDFromString function converts a string to a GUID, which is an array of type Byte.
    - `String As Variant` (required): A string expression that evaluates to a GUID in string form.
- `FollowHyperlink(Address As String, [SubAddress As String], [NewWindow As Boolean], [AddHistory As Boolean], [ExtraInfo As Variant], [Method As MsoExtraInfoMethod], [HeaderInfo As String])`  
  The FollowHyperlink method opens the document or webpage specified by a hyperlink address.
    - `Address As String` (required): A string expression that evaluates to a valid hyperlink address.
    - `SubAddress As String` (optional): A string expression that evaluates to a named location in the document specified by the _address_ argument. The default is a zero-length string (" ").
    - `NewWindow As Boolean` (optional): A Boolean value where True (1) opens the document in a new window and False (0) opens the document in the current window. The default is False.
    - `AddHistory As Boolean` (optional): A Boolean value where True adds the hyperlink to the History folder and False doesn't add the hyperlink to the History folder. The default is True.
    - `ExtraInfo As Variant` (optional): A string or an array of Byte data that specifies additional information for navigating to a hyperlink. For example, this argument may be used to specify a search parameter for an .asp or .idc file. In your web browser, the _extrainfo_ argument may appear after the hyperlink address, separated from the address by a question mark (?). You don't need to include the question mark when you specify the _extrainfo_ argument.
    - `Method As MsoExtraInfoMethod` (optional): An MsoExtraInfoMethod constant that specifies how the _extrainfo_ argument is attached.
    - `HeaderInfo As String` (optional): A string that specifies header information. By default, the _headerinfo_ argument is a zero-length string.
- `AddToFavorites()`  
  The AddToFavorites method adds a hyperlink address to the Favorites folder.
- `RefreshDatabaseWindow()`  
  The RefreshDatabaseWindow method updates the Database window after a database object has been created, deleted, or renamed.
- `RunCommand(Command As AcCommand)`  
  The RunCommand method runs a built-in command.
    - `Command As AcCommand` (required): An AcCommand constant that specifies the command to run.
- `HyperlinkPart(Hyperlink As Variant, [Part As AcHyperlinkPart]) As String`  
  The HyperlinkPart method returns information about data stored as a Hyperlink data type.
    - `Hyperlink As Variant` (required): The data stored in a Hyperlink field.
    - `Part As AcHyperlinkPart` (optional): An AcHyperlinkPart constant representing the information that you want returned by the HyperlinkPart method.
- `GetHiddenAttribute(ObjectType As AcObjectType, ObjectName As String) As Boolean`  
  The GetHiddenAttribute method returns the value of a hidden attribute of a Microsoft Access object in the object's Properties dialog box, available by selecting the object in the Database window and choosing Properties on the View menu.
    - `ObjectType As AcObjectType` (required): An AcObjectType constant that specifies the type of Access object.
    - `ObjectName As String` (required): The name of the Access object.
- `SetHiddenAttribute(ObjectType As AcObjectType, ObjectName As String, fHidden As Boolean)`  
  The SetHiddenAttribute method sets the hidden attribute of an Access object.
    - `ObjectType As AcObjectType` (required): An AcObjectType constant that specifies the type of Access object.
    - `ObjectName As String` (required): The name of the Access object.
    - `fHidden As Boolean` (required): True sets the hidden attribute, and False clears the attribute.
- `NewAccessProject(filepath As String, [Connect As Variant])`  
  Use the NewAccessProject method to create and open a new Microsoft Access project (.adp) as the current Access project in the Access window.
    - `filepath As String` (required): The name of the new Access project, including the path name and the file name extension.
    - `Connect As Variant` (optional): The connection string for the Access project. See the ADO ConnectionString property for details about this string.
- `OpenAccessProject(filepath As String, [Exclusive As Boolean])`  
  Use the OpenAccessProject method to open an existing Microsoft Access project (.adp) as the current Access project in the Access window.
    - `filepath As String` (required): The name of the existing Access project, including the path name and the file name extension.
    - `Exclusive As Boolean` (optional): Specifies whether you want to open the Access project in exclusive mode. The default value is False, which specifies that the Access project should be opened in shared mode.
- `CreateAccessProject(filepath As String, [Connect As Variant])`  
  Use the CreateAccessProject method to create a new Microsoft Access project (.adp) on disk.
    - `filepath As String` (required): A string expression that is the name of the new Access project, including the path name and the file name extension. If your network supports it, you can also specify a network path in the following form: \\Server\Share\Folder\Filename.adp
    - `Connect As Variant` (optional): A string expression that's the valid connection string for the Access project. See the ADO ConnectionString property for details about this string.
- `EuroConvert(Number As Double, SourceCurrency As String, TargetCurrency As String, [FullPrecision As Variant], [TriangulationPrecision As Variant]) As Double`  
  Use the EuroConvert function to convert a number to euro, or from euro to a participating currency. You can also use it to convert a number from one participating currency to another by using the euro as an intermediary (triangulation). The EuroConvert function uses fixed conversion rates established by the European Union.
    - `Number As Double` (required): The number you want to convert, or a reference to a field containing the number.
    - `SourceCurrency As String` (required): A string expression, or reference to a field containing the string, corresponding to the International Standards Organization (ISO) acronym for the currency you want to convert. Can be one of the ISO codes listed in the Remarks section.
    - `TargetCurrency As String` (required): A string expression, or reference to a field containing the string, corresponding to the ISO code of the currency to which you want to convert the number. For a list of ISO codes, see the Remarks section.
    - `FullPrecision As Variant` (optional): A Boolean value where True (1) ignores the currency-specific rounding rules (called display precision in the _sourcecurrency_ argument description) and uses the 6-significant-digit conversion factor with no follow-up rounding. False (0) uses the currency-specific rounding rules to display the result. If the parameter is omitted, the default value is False.
    - `TriangulationPrecision As Variant` (optional): An Integer value greater than or equal to 3 that specifies the number of significant digits in the calculation precision used for the intermediate euro value when converting between two national/regional currencies.
- `SetDefaultWorkgroupFile(Path As String)`  
  Sets the default workgroup file to the specified file.
    - `Path As String` (required): The full path and file name of the workgroup file to use as the default.
- `ConvertAccessProject(SourceFilename As String, DestinationFilename As String, DestinationFileFormat As AcFileFormat)`  
  Converts the specified Microsoft Access file from one version to another.
    - `SourceFilename As String` (required): The name of the Access file to convert. If a path isn't specified, Access looks for the file in the current directory.
    - `DestinationFilename As String` (required): The name of the file where Access saves the converted file. If a path isn't specified, Access saves the file in the current directory.
    - `DestinationFileFormat As AcFileFormat` (required): An AcFileFormat constant that specifies the format of the converted file.
- `OpenCurrentDatabase(filepath As String, [Exclusive As Boolean], [bstrPassword As String])`  
  Use the OpenCurrentDatabase method to open an existing Microsoft Access database as the current database.
    - `filepath As String` (required): The name of an existing database file, including the path name and the file name extension.
    - `Exclusive As Boolean` (optional): Specifies whether you want to open the database in exclusive mode. The default value is False, which specifies that the database should be opened in shared mode.
    - `bstrPassword As String` (optional): The password to open the specified database.
- `CompactRepair(SourceFile As String, DestinationFile As String, [LogFile As Boolean]) As Boolean`
- `ImportXML(DataSource As String, [ImportOptions As AcImportXMLOption])`  
  The ImportXML method allows developers to import XML data and/or schema information into Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 7.0 or later, or the Microsoft Access database engine.
    - `DataSource As String` (required): The name and path of the XML file to import.
    - `ImportOptions As AcImportXMLOption` (optional): An AcImportXMLOption constant that specifies the option to use when importing XML files. The default value is acStructureAndData.
- `TransformXML(DataSource As String, TransformSource As String, OutputTarget As String, [WellFormedXMLOutput As Boolean], [ScriptOption As AcTransformXMLScriptOption])`  
  Applies an Extensible Stylesheet Language (XSL) stylesheet to an XML data file and writes the resulting XML to an XML data file.
    - `DataSource As String` (required): The name and path of the XML file to import.
    - `TransformSource As String` (required): The name and path to the XSL file to apply to the XML data file.
    - `OutputTarget As String` (required): The file name and path for the resulting XML data file after applying the XSL file.
    - `WellFormedXMLOutput As Boolean` (optional): Setting this argument to True will create a well-formed XML file. Setting this argument to False will encode the resulting XML file in UTF-16 format. The default value is False.
    - `ScriptOption As AcTransformXMLScriptOption` (optional): An AcTransformXMLScriptOption constant that specifies the action taken if the XSL file contains scripting code. The default value is acPromptScript.
- `CreateAdditionalData() As _AdditionalData`  
  Creates an AdditionalData object that can be used to add additional tables and queries to the parent table that is being exported by the ExportXML method.
- `ExportXML(ObjectType As AcExportXMLObjectType, DataSource As String, [DataTarget As String], [SchemaTarget As String], [PresentationTarget As String], [ImageTarget As String], [Encoding As AcExportXMLEncoding], [OtherFlags As AcExportXMLOtherFlags], [WhereCondition As String], [AdditionalData As Variant])`  
  The ExportXML method allows developers to export XML data, schemas, and presentation information from Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 6.5 or later, or the Microsoft Access database engine.
    - `ObjectType As AcExportXMLObjectType` (required): An AcExportXMLObjectType that represents the type of AccessObject object to export.
    - `DataSource As String` (required): The name of the AccessObject object to export. The default is the currently open object of the type specified by the _ObjectType_ argument.
    - `DataTarget As String` (optional): The file name and path for the exported data. If this argument is omitted, data is not exported.
    - `SchemaTarget As String` (optional): The file name and path for the exported schema information. If this argument is omitted, schema information is not exported to a separate XML file.
    - `PresentationTarget As String` (optional): The file name and path for the exported presentation information. If this argument is omitted, presentation information is not exported.
    - `ImageTarget As String` (optional): The path for exported images. If this argument is omitted, images are not exported.
    - `Encoding As AcExportXMLEncoding` (optional): An AcExportXMLEncoding constant that specifies the text encoding to use for the exported XML. The default value is acUTF8.
    - `OtherFlags As AcExportXMLOtherFlags` (optional): A bit mask that specifies other behaviors associated with exporting to XML. Can be a combination of AcExportXMLOtherFlags constants.
    - `WhereCondition As String` (optional): Specifies a subset of records to be exported.
    - `AdditionalData As Variant` (optional): Specifies additional tables to export. This argument is ignored if the _OtherFlags_ argument is set to acLiveReportSource.
- `NewCurrentDatabase(filepath As String, [FileFormat As AcNewDatabaseFormat], [Template As Variant], [SiteAddress As String], [ListID As String])`  
  Creates a new Microsoft Access database.
    - `filepath As String` (required): A string expression that is the name of a new database file, including the path name and the file name extension. If your network supports it, you can also specify a network path in the following form: \\Server\Share\Folder\Filename
    - `FileFormat As AcNewDatabaseFormat` (optional): An AcNewDatabaseFormat constant that specifies the file format to use for the newly created database.
    - `Template As Variant` (optional): The name of the template to be used for the new database.
    - `SiteAddress As String` (optional): Uniform Resource Locator (URL) of the Windows SharePoint Services 3.0 site to link to.
    - `ListID As String` (optional): Globally Unique Identifer (GUID) or the name of the Windows SharePoint Services 3.0 list to link to.
- `PlainText(RichText As Variant, [Length As Variant]) As String`  
  Strips the rich text formatting from a string and returns an unformatted text string.
    - `RichText As Variant` (required): The rich text string that you want to remove the formatting of.
    - `Length As Variant` (optional): The number of characters to return.
- `HtmlEncode(PlainText As Variant, [Length As Variant]) As String`  
  Converts a string to an HTML-encoded string.
    - `PlainText As Variant` (required): The text string to encode.
    - `Length As Variant` (optional): The number of characters to return.
- `LoadCustomUI(CustomUIName As String, CustomUIXML As String)`  
  Loads XML markup that represents a customized ribbon.
    - `CustomUIName As String` (required): The name that will be used to identify the customized ribbon.
    - `CustomUIXML As String` (required): The XML markup code that defines the customized ribbon.
- `ExportNavigationPane(Path As String)`  
  Saves the current configuration of the navigation pane to an XML file.
    - `Path As String` (required): The path and name of the XML file that you want to save the configuration of the navigation pane to.
- `ImportNavigationPane(Path As String, [fAppendOnly As Boolean])`  
  Loads a saved navigation pane configuration from disk.
    - `Path As String` (required): The path and name of the XML file that contains the navigation pane configuration to load.
    - `fAppendOnly As Boolean` (optional): Set to True to append the imported categories to the existing categories. The default value is False.
- `ColumnHistory(TableName As String, ColumnName As String, queryString As String) As String`  
  Gets the history of values that have been stored in a Memo field.
    - `TableName As String` (required): The name of the table that contains the Append Only field.
    - `ColumnName As String` (required): The name of the field to display the history for.
    - `queryString As String` (required): A String used to locate the record. It's like the WHERE clause in an SQL statement, but without the word WHERE.
- `CreateControl(FormName As String, ControlType As AcControlType, [Section As AcSection], [Parent As Variant], [ColumnName As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As _Control`  
  The CreateControl method creates a control on a specified open form. For example, suppose you are building a custom wizard that allows users to easily construct a particular form. Use the CreateControl method in your wizard to add the appropriate controls to the form.
    - `FormName As String` (required): The name of the open form or report on which you want to create the control.
    - `ControlType As AcControlType` (required): An AcControlType constant that represents the type of control that you want to create.
    - `Section As AcSection` (optional): An AcSection constant that identifies the section that will contain the new control.
    - `Parent As Variant` (optional): The name of the parent control of an attached control. For controls that have no parent control, use a zero-length string for this argument or omit it.
    - `ColumnName As Variant` (optional): The name of the field to which the control will be bound if it is to be a data-bound control.
- `CreateReportControl(ReportName As String, ControlType As AcControlType, [Section As AcSection], [Parent As Variant], [ColumnName As Variant], [Left As Variant], [Top As Variant], [Width As Variant], [Height As Variant]) As _Control`  
  The CreateReportControl method creates a control on a specified open report. For more information, see the CreateControl method.
    - `ReportName As String` (required): The name of the open report on which you want to create the control.
    - `ControlType As AcControlType` (required): An AcControlType constant that represents the type of control that you want to create.
    - `Section As AcSection` (optional): An AcSection constant that identifies the section that will contain the new control.
    - `Parent As Variant` (optional): A string expression that identifies the name of the parent control of an attached control. For controls that have no parent control, use a zero-length string for this argument or omit it.
    - `ColumnName As Variant` (optional): The name of the field to which the control will be bound if it is to be a data-bound control.
- `SaveAsAXL(ObjectType As AcObjectType, ObjectName As String, FileName As String)`  
  Exports the specified object to an Application XML (AXL) file.
    - `ObjectType As AcObjectType` (required): Specifies the type of object to export.
    - `ObjectName As String` (required): Specifies the name of the object to export.
    - `FileName As String` (required): Specifies the full path and file name of the AXL file to create.
- `LoadFromAXL(ObjectType As AcObjectType, ObjectName As String, FileName As String)`  
  Imports the object defined in an Application XML (AXL) file into the database.
    - `ObjectType As AcObjectType` (required): Specifies the type of object to create.
    - `ObjectName As String` (required): Specifies the name of the object.
    - `FileName As String` (required): Specifies the full path and file name of the AXL file to import.
- `SaveAsTemplate(Path As String, Title As String, IconPath As String, CoreTable As String, Category As String, [PreviewPath As Variant], [Description As Variant], [InstantiationForm As Variant], [ApplicationPart As Variant], [IncludeData As Variant], [Variation As Variant])`  
  Converts an existing Microsoft Access database file to a database template (*.accdt) format file.
    - `Path As String` (required): The full path and file name of the database template to create.
    - `Title As String` (required): The name of the database that is created when the user instantiates the template.
    - `IconPath As String` (required): An image file to be used as an icon for the database template.
    - `CoreTable As String` (required): The table that contains the data that users would most want to create a relationship with when they instantiate the template. The _ApplicationPart_ argument must be set to True if you use this argument.
    - `Category As String` (required): The template category under which the database template will appear on the Available Templates page.
    - `PreviewPath As Variant` (optional): An image file to be used as a preview for the database template on the Available Templates page.
    - `Description As Variant` (optional): A description to be displayed when the user selects the database template on the Available Templates page.
    - `InstantiationForm As Variant` (optional): Specifies the name of the form to be displayed when the template is instantiated.
    - `ApplicationPart As Variant` (optional): Specifies whether the template will be displayed when the user chooses Application Parts in the Templates group of the Create ribbon tab. Set to True to display the template when the user chooses Application Parts.
    - `IncludeData As Variant` (optional): Specifies whether the table data is included in the template. Set to True to include the table data.
- `InstantiateTemplate(Path As String)`  
  Opens a new database and applies the specified template.
    - `Path As String` (required): The full path and file name of the template to apply to the new database.
- `CurrentWebUser(DisplayOption As AcWebUserDisplay) As Variant`  
  Gets information about the current user of a web database on Microsoft SharePoint Foundation 2010 and later.
    - `DisplayOption As AcWebUserDisplay` (required): Specifies the type of information to return about the user.
- `CurrentWebUserGroups(DisplayOption As AcWebUserGroupsDisplay) As Variant`  
  Gets the collection of Microsoft SharePoint Foundation groups of which the user is a member.
    - `DisplayOption As AcWebUserGroupsDisplay` (required): Specifies the type of information to return about the user's groups.
- `IsCurrentWebUserInGroup(GroupNameOrID As Variant) As Boolean`  
  Gets whether or not the current user of a web database is a member of the specified Microsoft SharePoint Foundation 2010 group.
    - `GroupNameOrID As Variant` (required): The name or identifier (ID) of the group.
- `DirtyObject(ObjectType As AcObjectType, ObjectName As String)`  
  Marks a form or report as dirty.
    - `ObjectType As AcObjectType` (required): Specifies the type of object to mark as dirty. This argument should be set to acForm or acReport.
    - `ObjectName As String` (required): Specifies the name of the object to mark as dirty.
