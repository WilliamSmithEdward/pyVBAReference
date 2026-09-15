# DoCmd

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {C547E760-9658-101B-81EE-00AA004750E2}  

Use the methods of the DoCmd object to run Microsoft Office Access actions from Visual Basic. An action performs tasks such as closing windows, opening forms, and setting the value of controls.

**Remarks:** For example, you can use the OpenForm method of the DoCmd object to open a form, or use the Hourglass method to change the mouse pointer to an hourglass icon. Most of the methods of the DoCmd object have arguments; some are required, while others are optional. If you omit optional arguments, the arguments assume the default values for the particular method. For example, the OpenForm method uses seven arguments, but only the first argument, FormName, is required. The following example shows how you can open the Employees form in the current database. Only employees with the title Sales Representative are included. The DoCmd object doesn't support methods corresponding to the following actions: - MsgBox. Use the MsgBox function. - RunApp. Use the Shell function to run another application. - RunCode. Run the function directly in Visual Basic. - SendKeys. Use the SendKeys statement. - SetValue. Set the value directly in Visual Basic. - StopAllMacros. - StopMacro.

**Example:**

```vba
Sub ShowNewRecord()
 DoCmd.OpenForm "Employees", acNormal
 DoCmd.GoToRecord , , acNewRec
End Sub
```

## Methods (66)

- `AddMenu(MenuName As Variant, MenuMacroName As Variant, [StatusBarText As Variant])`  
  The AddMenu method carries out the AddMenu action in Visual Basic.
    - `MenuName As Variant` (required): A string expression that's the valid name of a drop-down menu to add to the custom menu bar or global menu bar. To create an access key so that you can use the keyboard to choose the menu, type an ampersand (&) before the letter you want to be the access key. This letter will be underlined in the menu name on the menu bar.
    - `MenuMacroName As Variant` (required): A string expression that's the valid name of the macro group that contains the macros for the menu's commands. This is a required argument.
    - `StatusBarText As Variant` (optional): A string expression that's the text to display in the status bar when the menu is selected.
- `Beep()`  
  The Beep method carries out the Beep action in Visual Basic.
- `CancelEvent()`  
  The CancelEvent method carries out the CancelEvent action in Visual Basic.
- `Close([ObjectType As AcObjectType], [ObjectName As Variant], [Save As AcCloseSave])`  
  The Close method carries out the Close action in Visual Basic.
    - `ObjectType As AcObjectType` (optional): An AcObjectType constant that represents the type of object to close.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument.
    - `Save As AcCloseSave` (optional): An AcCloseSave constant that specifies whether to save changes to the object. The default value is acSavePrompt.
- `CopyObject([DestinationDatabase As Variant], [NewName As Variant], [SourceObjectType As AcObjectType], [SourceObjectName As Variant])`  
  The CopyObject method carries out the CopyObject action in Visual Basic.
    - `DestinationDatabase As Variant` (optional): A string expression that's the valid path and file name for the database that you want to copy the object into. To select the current database, leave this argument blank. NOTE: In a Microsoft Access project (.adp), you must leave the DestinationDatabase argument blank. If you execute Visual Basic code containing the CopyObject method in a library database and leave this argument blank, Access copies the object into the library database.
    - `NewName As Variant` (optional): A string expression that's the new name for the object that you want to copy. To use the same name if you are copying into another database, leave this argument blank.
    - `SourceObjectType As AcObjectType` (optional): An AcObjectType constant that represents the type of object to copy.
    - `SourceObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the SourceObjectType argument. If you run Visual Basic code containing the CopyObject method in a library database, Access looks for the object with this name first in the library database, and then in the current database.
- `DoMenuItem(MenuBar As Variant, MenuName As Variant, Command As Variant, [Subcommand As Variant], [Version As Variant])`  
  Displays the appropriate menu or toolbar command for Microsoft Access.
    - `MenuBar As Variant` (required): Use the intrinsic constant acFormBar for the menu bar in Form view. For other views, use the number of the view in the MenuBar argument list, as shown in the Macro window in previous versions of Microsoft Access (count down the list, starting from 0).
    - `MenuName As Variant` (required): Use one of the following intrinsic constants:<ul><li><p><b>acFile</b></p></li><li><p><b>acEditMenu</b></p></li><li><p><b>acRecordsMenu</b></p></li></ul><p>Use <b>acRecordsMenu</b> only for the Form view menu bar in Access version 2.0 and Access 95 databases. For other menus, use the number of the menu in the MenuName argument list, as shown in the Macro window in previous versions of Access (count down the list, starting from 0).</p>
    - `Command As Variant` (required): Use one of the following intrinsic constants:<ul><li><p><b>acNew</b></p></li><li><p><b>acSaveForm</b></p></li><li><p><b>acSaveFormAs</b></p></li><li><p><b>acSaveRecord</b></p></li><li><p><b>acUndo</b></p></li><li><p><b>acCut</b></p></li><li><p><b>acCopy</b></p></li><li><p><b>acPaste</b></p></li><li><p><b>acDelete</b></p></li><li><p><b>acSelectRecord</b></p></li><li><p><b>acSelectAllRecords</b></p></li><li><p><b>acObjectRefresh</b></p></li></ul><p>For other commands, use the number of the command in the Command argument list, as shown in the Macro window in previous versions of Access (count down the list, starting from 0).</p>
    - `Subcommand As Variant` (optional): Use one of the following intrinsic constants:<ul><li><p><b>acObjectVerb</b></p></li><li><p><b>acObjectUpdate</b></p></li></ul><p>The <b>acObjectVerb</b> constant represents the first command on the submenu of the <b>Object</b> command on the <b>Edit</b> menu. The type of object determines the first command on the submenu. For example, this command is Edit for a Paintbrush object that can be edited.</p> <p>For other commands on submenus, use the number of the subcommand in the Subcommand argument list, as shown in the Macro window in previous versions of Access (count down the list, starting from 0).</p>
    - `Version As Variant` (optional): Use the intrinsic constant acMenuVer70 for code written for Access 95 databases, the intrinsic constant acMenuVer20 for code written for Access version 2.0 databases, and the intrinsic constant acMenuVer1X for code written for Access version 1.x databases. This argument is available only in Visual Basic. NOTE: The default for this argument is acMenuVer1X, so that any code written for Access version 1.x databases will run unchanged. If you are writing code for a Access 95 or version 2.0 database and want to use the Access 95 or version 2.0 menu commands with the DoMenuItem method, you must set this argument to acMenuVer70 or acMenuVer20. Also, when you are counting down the lists for the MenuBar, MenuName, Command, and Subcommand action arguments in the Macro window to get the numbers to use for the arguments in the DoMenuItem method, you must use the Access 95 lists if the Version argument is acMenuVer70, the Access version 2.0 lists if the Version argument is Version, and the Access version 1.x lists if Version is acMenuVer1X (or blank). NOTE: There is no acMenuVer80 setting for this argument. You can't use the DoMenuItem method to display Access commands (although existing DoMenuItem methods in Visual Basic code will still work). Use the RunCommand method instead.
- `Echo(EchoOn As Variant, [StatusBarText As Variant])`  
  Carries out the Echo action in Visual Basic.
    - `EchoOn As Variant` (required): Use True to turn echo on and False to turn it off.
    - `StatusBarText As Variant` (optional): A string expression indicating the text that appears in the status bar.
- `FindNext()`  
  The FindNext method carries out the FindNext action in Visual Basic.
- `FindRecord(FindWhat As Variant, [Match As AcFindMatch], [MatchCase As Variant], [Search As AcSearchDirection], [SearchAsFormatted As Variant], [OnlyCurrentField As AcFindField], [FindFirst As Variant])`  
  The FindRecord method carries out the FindRecord action in Visual Basic.
    - `FindWhat As Variant` (required): An expression that evaluates to text, a number, or a date. The expression contains the data to search for.
    - `Match As AcFindMatch` (optional): An AcFindMatch constant that specifies where to search for the match. The default value is acEntire.
    - `MatchCase As Variant` (optional): Use True for a case-sensitive search and False for a search that's not case-sensitive. If you leave this argument blank, the default (False) is assumed.
    - `Search As AcSearchDirection` (optional): An AcSearchDirection constant that specifies the direction to search. The default value is acSearchAll.
    - `SearchAsFormatted As Variant` (optional): Use True to search for data as it's formatted and False to search for data as it's stored in the database. If you leave this argument blank, the default (False) is assumed.
    - `OnlyCurrentField As AcFindField` (optional): An AcFindField constant that specifies whether to search all fields, or only the current field. The default value is acCurrent.
    - `FindFirst As Variant` (optional): Use True to start the search at the first record. Use False to start the search at the record following the current record. If you leave this argument blank, the default (True) is assumed.
- `GoToControl(ControlName As Variant)`  
  The GoToControl method performs the GoToControl action in Visual Basic.
    - `ControlName As Variant` (required): A string expression that is the name of a control on the active form or datasheet.
- `GoToPage(PageNumber As Variant, [Right As Variant], [Down As Variant])`  
  Carries out the GoToPage action in Visual Basic.
    - `PageNumber As Variant` (required): A numeric expression that's a valid page number for the active form. If you leave this argument blank, the focus stays on the current page. Use the Right and Down arguments to display the part of the page that you want to see.
    - `Right As Variant` (optional): A numeric expression that's a valid horizontal offset for the page.
    - `Down As Variant` (optional): A numeric expression that's a valid vertical offset for the page.
- `GoToRecord([ObjectType As AcDataObjectType], [ObjectName As Variant], [Record As AcRecord], [Offset As Variant])`  
  The GoToRecord method carries out the GoToRecord action in Visual Basic.
    - `ObjectType As AcDataObjectType` (optional): An AcDataObjectType constant that specifies the type of object that contains the record that you want to make current.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument.
    - `Record As AcRecord` (optional): An AcRecord constant that specifies the record to make the current record. The default value is acNext.
    - `Offset As Variant` (optional): A numeric expression that represents the number of records to move forward or backward if you specify acNext or acPrevious for the Record argument, or the record to move to if you specify acGoTo for the Record argument. The expression must result in a valid record number.
- `Hourglass(HourglassOn As Variant)`  
  The Hourglass method carries out the Hourglass action in Visual Basic.
    - `HourglassOn As Variant` (required): Use True (1) to display the hourglass icon (or another icon you've chosen). Use False (0) to display the normal mouse pointer.
- `Maximize()`  
  The Maximize method carries out the Maximize action in Visual Basic.
- `Minimize()`  
  The Minimize method carries out the Minimize action in Visual Basic.
- `MoveSize([Right As Variant], [Down As Variant], [Width As Variant], [Height As Variant])`  
  The MoveSize method carries out the MoveSize action in Visual Basic.
    - `Right As Variant` (optional): The new horizontal position of the window's upper-left corner, measured from the left edge of its containing window.
    - `Down As Variant` (optional): The new vertical position of the window's upper-left corner, measured from the top edge of its containing window.
    - `Width As Variant` (optional): The window's new width.
    - `Height As Variant` (optional): The window's new height.
- `OpenForm(FormName As Variant, [View As AcFormView], [FilterName As Variant], [WhereCondition As Variant], [DataMode As AcFormOpenDataMode], [WindowMode As AcWindowMode], [OpenArgs As Variant])`  
  The OpenForm method carries out the OpenForm action in Visual Basic.
    - `FormName As Variant` (required): A string expression that's the valid name of a form in the current database. If you execute Visual Basic code containing the OpenForm method in a library database, Access looks for the form with this name first in the library database, and then in the current database.
    - `View As AcFormView` (optional): An AcFormView constant that specifies the view in which the form will open. The default value is acNormal.
    - `FilterName As Variant` (optional): A string expression that's the valid name of a query in the current database.
    - `WhereCondition As Variant` (optional): A string expression that's a valid SQL WHERE clause without the word WHERE.
    - `DataMode As AcFormOpenDataMode` (optional): An AcFormOpenDataMode constant that specifies the data entry mode for the form. This applies only to forms opened in Form view or Datasheet view. The default value is acFormPropertySettings.
    - `WindowMode As AcWindowMode` (optional): An AcWindowMode constant that specifies the window mode in which the form opens. The default value is acWindowNormal.
    - `OpenArgs As Variant` (optional): A string expression. This expression is used to set the form's OpenArgs property. This setting can then be used by code in a form module, such as the Open event procedure. The OpenArgs property can also be referred to in macros and expressions. For example, suppose that the form that you open is a continuous-form list of clients. If you want the focus to move to a specific client record when the form opens, you can specify the client name with the OpenArgs argument, and then use the FindRecord method to move the focus to the record for the client with the specified name.
- `OpenQuery(QueryName As Variant, [View As AcView], [DataMode As AcOpenDataMode])`  
  The OpenQuery method carries out the OpenQuery action in Visual Basic.
    - `QueryName As Variant` (required): A string expression that's the valid name of a query in the current database. If you execute Visual Basic code containing the OpenQuery method in a library database, Microsoft Access looks for the query with this name first in the library database, and then in the current database.
    - `View As AcView` (optional): An AcView constant that specifies the view in which the query will open. The default value is acViewNormal.
    - `DataMode As AcOpenDataMode` (optional): An AcOpenDataMode constant that specifies the data entry mode for the query. The default value is acEdit.
- `OpenTable(TableName As Variant, [View As AcView], [DataMode As AcOpenDataMode])`  
  The OpenTable method carries out the OpenTable action in Visual Basic.
    - `TableName As Variant` (required): A string expression that's the valid name of a table in the current database. If you execute Visual Basic code containing the OpenTable method in a library database, Microsoft Access looks for the table with this name first in the library database, and then in the current database.
    - `View As AcView` (optional): An AcView constant that specifies the view in which the table will open. The default value is acViewNormal.
    - `DataMode As AcOpenDataMode` (optional): An AcOpenDataMode constant that specifies the data entry mode for the table. The default value is acEdit.
- `PrintOut([PrintRange As AcPrintRange], [PageFrom As Variant], [PageTo As Variant], [PrintQuality As AcPrintQuality], [Copies As Variant], [CollateCopies As Variant])`  
  The PrintOut method carries out the PrintOut action in Visual Basic.
    - `PrintRange As AcPrintRange` (optional): An AcPrintRange constant that specifies the range to print. The default value is acPrintAll.
    - `PageFrom As Variant` (optional): The first page to print. A numeric expression that's a valid page number in the active form or datasheet. This argument is required if you specify acPages for the PrintRange argument.
    - `PageTo As Variant` (optional): The last page to print. A numeric expression that's a valid page number in the active form or datasheet. This argument is required if you specify acPages for the PrintRange argument.
    - `PrintQuality As AcPrintQuality` (optional): An AcPrintQuality constant that specifies the print quality. The default value is acHigh.
    - `Copies As Variant` (optional): The number of copies to print. If you leave this argument blank, the default (1) is assumed.
    - `CollateCopies As Variant` (optional): Use True (1) to collate copies and False (0) to print without collating. If you leave this argument blank, the default (True) is assumed.
- `Quit([Options As AcQuitOption])`  
  The Quit method quits Microsoft Access. You can select one of several options for saving a database object before quitting.
    - `Options As AcQuitOption` (optional): An AcQuitOption constant that indicates the action to take when quitting Access. The default value is acQuitSaveAll.
- `Requery([ControlName As Variant])`  
  Carries out the Requery action in Visual Basic.
    - `ControlName As Variant` (optional): A string expression that's the name of a control on the active object.
- `RepaintObject([ObjectType As AcObjectType], [ObjectName As Variant])`  
  The RepaintObject method carries out the RepaintObject action in Visual Basic.
    - `ObjectType As AcObjectType` (optional): An AcObjectType constant that specifies the type of object to repaint.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument.
- `Rename(NewName As Variant, [ObjectType As AcObjectType], [OldName As Variant])`  
  The Rename method carries out the Rename action in Visual Basic.
    - `NewName As Variant` (required): A string expression that's the new name for the object that you want to rename. The name must follow the object-naming rules for Microsoft Access objects.
    - `ObjectType As AcObjectType` (optional): An AcObjectType constant that specifies the type of object to rename. The default value is acDefault.
    - `OldName As Variant` (optional): A string expression that's the valid name of an object of the type specified by the ObjectType argument. If you execute Visual Basic code containing the Rename method in a library database, Access looks for the object with this name first in the library database, and then in the current database.
- `Restore()`  
  The Restore method carries out the Restore action in Visual Basic.
- `RunMacro(MacroName As Variant, [RepeatCount As Variant], [RepeatExpression As Variant])`  
  The RunMacro method carries out the RunMacro action in Visual Basic.
    - `MacroName As Variant` (required): A string expression that's the valid name of a macro in the current database. If you run Visual Basic code containing the RunMacro method in a library database, Microsoft Access looks for the macro with this name in the library database and doesn't look for it in the current database.
    - `RepeatCount As Variant` (optional): A numeric expression that evaluates to an integer, which is the number of times the macro will run.
    - `RepeatExpression As Variant` (optional): A numeric expression that's evaluated each time the macro runs. When it evaluates to False (0), the macro stops running.
- `RunSQL(SQLStatement As Variant, [UseTransaction As Variant])`  
  The RunSQL method carries out the RunSQL action in Visual Basic.
    - `SQLStatement As Variant` (required): A string expression that's a valid SQL statement for an action query or a data-definition query. It uses an INSERT INTO, DELETE, SELECT...INTO, UPDATE, CREATE TABLE, ALTER TABLE, DROP TABLE, CREATE INDEX, or DROP INDEX statement. Include an IN clause if you want to access another database.
    - `UseTransaction As Variant` (optional): Use True (1) to include this query in a transaction. Use False (0) if you don't want to use a transaction. If you leave this argument blank, the default (True) is assumed.
- `SelectObject(ObjectType As AcObjectType, [ObjectName As Variant], [InDatabaseWindow As Variant])`  
  The SelectObject method carries out the SelectObject action in Visual Basic.
    - `ObjectType As AcObjectType` (required): An AcObjectType constant that specifies the type of object that you want to select.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument. This is a required argument, unless you specify True (1) for the InNavigationPane argument.
- `SetWarnings(WarningsOn As Variant)`  
  The SetWarnings method carries out the SetWarnings action in Visual Basic.
    - `WarningsOn As Variant` (required): Use True (1) to turn on the display of system messages and False (0) to turn it off.
- `ShowAllRecords()`  
  The ShowAllRecords method carries out the ShowAllRecords action in Visual Basic.
- `TransferDatabase([TransferType As AcDataTransferType], [DatabaseType As Variant], [DatabaseName As Variant], [ObjectType As AcObjectType], [Source As Variant], [Destination As Variant], [StructureOnly As Variant], [StoreLogin As Variant])`  
  The TransferDatabase method carries out the TransferDatabase action in Visual Basic.
    - `TransferType As AcDataTransferType` (optional): The type of transfer you want to make.
    - `DatabaseType As Variant` (optional): A string expression that's the name of one of the types of databases that you can use to import, export, or link data. The DatabaseType parameter is required for exporting and link data actions but not required for importing actions. The types or databases are:<ul><li><p>Microsoft Access (default) </p></li><li><p>Jet 2.x</p></li><li><p>Jet 3.x</p></li><li><p>dBase III</p></li><li><p>dBase IV</p></li><li><p>dBase 5.0</p></li><li><p>Paradox 3.x</p></li><li><p>Paradox 4.x</p></li><li><p>Paradox 5.x</p></li><li><p>Paradox 7.x</p></li><li><p>ODBC Database</p></li><li><p>WSS (SharePoint)</p></li></ul>
    - `DatabaseName As Variant` (optional): A string expression that's the full name, including the path (for WSS, Windows SharePoint Services, the URL), of the database that you want to use to import, export, or link data.
    - `ObjectType As AcObjectType` (optional): The type of object to import or export.
    - `Source As Variant` (optional): A string expression that's the name of the object whose data you want to import, export, or link.
    - `Destination As Variant` (optional): A string expression that's the name of the imported, exported, or linked object in the destination database.
    - `StructureOnly As Variant` (optional): Use True (1) to import or export only the structure of a database table. Use False (0) to import or export the structure of the table and its data. If you leave this argument blank, the default (False) is assumed.
    - `StoreLogin As Variant` (optional): Use True to store the sign-in identification (ID) and password for an ODBC database in the connection string for a linked table from the database. If you do this, you don't have to sign in each time you open the table. Use False if you don't want to store the sign-in ID and password. If you leave this argument blank, the default (False) is assumed. This argument is available only in Visual Basic.
- `TransferSpreadsheet([TransferType As AcDataTransferType], [SpreadsheetType As AcSpreadSheetType], [TableName As Variant], [FileName As Variant], [HasFieldNames As Variant], [Range As Variant], [UseOA As Variant])`  
  The TransferSpreadsheet method carries out the TransferSpreadsheet action in Visual Basic.
    - `TransferType As AcDataTransferType` (optional): The type of transfer that you want to make. The default value is acImport.
    - `SpreadsheetType As AcSpreadSheetType` (optional): The type of spreadsheet to import from, export to, or link to.
    - `TableName As Variant` (optional): A string expression that is the name of the Office Access table that you want to import spreadsheet data into, export spreadsheet data from, or link spreadsheet data to, or the Access select query whose results you want to export to a spreadsheet.
    - `FileName As Variant` (optional): A string expression that's the file name and path of the spreadsheet that you want to import from, export to, or link to.
    - `HasFieldNames As Variant` (optional): Use True (1) to use the first row of the spreadsheet as field names when importing or linking. Use False (0) to treat the first row of the spreadsheet as normal data. If you leave this argument blank, the default (False) is assumed. When you export Access table or select query data to a spreadsheet, the field names are inserted into the first row of the spreadsheet no matter what you enter for this argument.
    - `Range As Variant` (optional): A string expression that's a valid range of cells or the name of a range in the spreadsheet. This argument applies only to importing. Leave this argument blank to import the entire spreadsheet. When you export to a spreadsheet, you must leave this argument blank. If you enter a range, the export will fail.
    - `UseOA As Variant` (optional): This argument is not supported.
- `TransferText([TransferType As AcTextTransferType], [SpecificationName As Variant], [TableName As Variant], [FileName As Variant], [HasFieldNames As Variant], [HTMLTableName As Variant], [CodePage As Variant])`  
  The TransferText method carries out the TransferText action in Visual Basic.
    - `TransferType As AcTextTransferType` (optional): The type of transfer you want to make. You can import data from, export data to, or link to data in delimited or fixed-width text files or HTML files. The default value is acImportDelim. Only acImportDelim, acImportFixed, acExportDelim, acExportFixed, or acExportMerge transfer types are supported in a Microsoft Access project (.adp).
    - `SpecificationName As Variant` (optional): A string expression that's the name of an import or export specification you've created and saved in the current database. For a fixed-width text file, you must either specify an argument or use a schema.ini file, which must be stored in the same folder as the imported, linked, or exported text file. To create a schema file, you can use the text import/export wizard to create the file. For delimited text files and Microsoft Word mail merge data files, you can leave this argument blank to select the default import/export specifications.
    - `TableName As Variant` (optional): A string expression that's the name of the Access table you want to import text data to, export text data from, or link text data to, or the Access query whose results you want to export to a text file.
    - `FileName As Variant` (optional): A string expression that's the full name, including the path, of the text file you want to import from, export to, or link to.
    - `HasFieldNames As Variant` (optional): Use True (1) to use the first row of the text file as field names when importing, exporting, or linking. Use False (0) to treat the first row of the text file as normal data. If you leave this argument blank, the default (False) is assumed. This argument is ignored for Microsoft Word mail merge data files, which must always contain the field names in the first row.
    - `HTMLTableName As Variant` (optional): A string expression that's the name of the table or list in the HTML file that you want to import or link. This argument is ignored unless the TransferType argument is set to acImportHTML or acLinkHTML. If you leave this argument blank, the first table or list in the HTML file is imported or linked. The name of the table or list in the HTML file is determined by the text specified by the CAPTION tag, if there's a CAPTION tag. If there's no CAPTION tag, the name is determined by the text specified by the TITLE tag. If more than one table or list has the same name, Access distinguishes them by adding a number to the end of each table or list name; for example, Employees1 and Employees2.
- `DeleteObject([ObjectType As AcObjectType], [ObjectName As Variant])`  
  The DeleteObject method carries out the DeleteObject action in Visual Basic.
    - `ObjectType As AcObjectType` (optional): An AcObjectType constant that represents the type of object to delete.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument. If you run Visual Basic code containing the DeleteObject method in a library database, Microsoft Access looks for the object with this name first in the library database, and then in the current database.
- `OpenModule([ModuleName As Variant], [ProcedureName As Variant])`  
  The OpenModule method carries out the OpenModule action in Visual Basic.
    - `ModuleName As Variant` (optional): A string expression that's the valid name of the Visual Basic module that you want to open. If you leave this argument blank, Microsoft Access searches all the standard modules in the database for the procedure that you selected with the ProcedureName argument and opens the module containing the procedure to that procedure. If you execute Visual Basic code containing the OpenModule method in a library database, Access looks for the module with this name first in the library database, and then in the current database.
    - `ProcedureName As Variant` (optional): A string expression that's the valid name for the procedure that you want to open the module to. If you leave this argument blank, the module opens to the Declarations section.
- `SendObject([ObjectType As AcSendObjectType], [ObjectName As Variant], [OutputFormat As Variant], [To As Variant], [Cc As Variant], [Bcc As Variant], [Subject As Variant], [MessageText As Variant], [EditMessage As Variant], [TemplateFile As Variant])`  
  The SendObject method carries out the SendObject action in Visual Basic.
    - `ObjectType As AcSendObjectType` (optional): An AcSendObjectType constant that specifies the type of object to send.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument. If you want to include the active object in the mail message, specify the object's type with the ObjectType argument and leave this argument blank. If you leave both the ObjectType and ObjectName arguments blank (the default constant, acSendNoObject, is assumed for the ObjectType argument), Microsoft Access sends a message to the electronic mail application without an included database object. If you run Visual Basic code containing the SendObject method in a library database, Access looks for the object with this name first in the library database, and then in the current database.
    - `OutputFormat As Variant` (optional): A constant that specifies the format in which to send the object. Possible values include acFormatHTML, acFormatRTF, acFormatSNP, acFormatTXT, acFormatXLS, acFormatXLSB, acFormatXLSX, acFormatXPS, and acFormatPDF.
    - `To As Variant` (optional): A string expression that lists the recipients whose names you want to put on the To line in the mail message. Separate the recipient names that you specify in this argument and in the Cc and Bcc arguments with a semicolon (;), or with the list separator set on the Number tab of the Regional Settings Properties dialog box in the Windows Control Panel. If the recipient names aren't recognized by the mail application, the message isn't sent and an error occurs. If you leave this argument blank, Microsoft Access prompts you for the recipients.
    - `Cc As Variant` (optional): A string expression that lists the recipients whose names you want to put on the Cc line in the mail message. If you leave this argument blank, the Cc line in the mail message is blank.
    - `Bcc As Variant` (optional): A string expression that lists the recipients whose names you want to put on the Bcc line in the mail message. If you leave this argument blank, the Bcc line in the mail message is blank.
    - `Subject As Variant` (optional): A string expression containing the text that you want to put on the Subject line in the mail message. If you leave this argument blank, the Subject line in the mail message is blank.
    - `MessageText As Variant` (optional): A string expression containing the text that you want to include in the body of the mail message, after the object. If you leave this argument blank, the object is all that's included in the body of the mail message.
    - `EditMessage As Variant` (optional): Use True (1) to open the electronic mail application immediately with the message loaded, so the message can be edited. Use False (0) to send the message without editing it. If you leave this argument blank, the default (True) is assumed.
    - `TemplateFile As Variant` (optional): A string expression that's the full name, including the path, of the file that you want to use as a template for an HTML file.
- `ShowToolbar(ToolbarName As Variant, [Show As AcShowToolbar])`  
  The ShowToolbar method carries out the ShowToolbar action in Visual Basic.
    - `ToolbarName As Variant` (required): A string expression that's the valid name of a custom toolbar you've created. If you run Visual Basic code containing the ShowToolbar method in a library database, Microsoft Access looks for the toolbar with this name first in the library database, and then in the current database.
    - `Show As AcShowToolbar` (optional): An AcShowToolbar constant that specifies whether to display or hide the toolbar and in which views to display or hide it. The default value is acToolbarYes.
- `Save([ObjectType As AcObjectType], [ObjectName As Variant])`  
  The Save method carries out the Save action in Visual Basic.
    - `ObjectType As AcObjectType` (optional): An AcObjectType constant that specifies the type of object that you want to save.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument.
- `SetMenuItem(MenuIndex As Variant, [CommandIndex As Variant], [SubcommandIndex As Variant], [Flag As Variant])`  
  The SetMenuItem method carries out the SetMenuItem action in Visual Basic.
    - `MenuIndex As Variant` (required): An integer, counting from 0, that is the valid index of a menu on the custom menu bar or global menu bar for the active window, as defined in the menu bar macro for the custom menu bar or global menu bar. If you select a menu with this argument and leave the CommandIndex and SubcommandIndex arguments blank (or set them to 1), you can enable or disable the menu name itself. You can't, however, check or uncheck a menu name (Microsoft Access ignores the acMenuCheck and acMenuUncheck settings for the Flag argument for menu names).
    - `CommandIndex As Variant` (optional): An integer, counting from 0, that's the valid index of a command on the menu selected by the MenuIndex argument, as defined in the macro group that defines the selected menu for the custom menu bar or global menu bar for the active window.
    - `SubcommandIndex As Variant` (optional): An integer, counting from 0, that's the valid index of a subcommand in the submenu selected by the CommandIndex argument, as defined in the macro group that defines the selected submenu for the custom menu bar or global menu bar for the active window.
    - `Flag As Variant` (optional): The state you want to set the command or subcommand to. Can be one of the following constants:<ul><li><b>acMenuCheck</b></li><li><b>acMenuGray</b></li><li><b>acMenuUncheck</b></li><li><b>acMenuUngray</b> (default)</li></ul>
- `RunCommand(Command As AcCommand)`  
  The RunCommand method runs a built-in command.
    - `Command As AcCommand` (required): An AcCommand constant that specifies the command to run.
- `OpenDataAccessPage(DataAccessPageName As Variant, [View As AcDataAccessPageView])`  
  The OpenDataAccessPage method carries out the OpenDataAccessPage action in Visual Basic.
    - `DataAccessPageName As Variant` (required): A string expression that's the valid name of a data access page in the current database. If you execute Visual Basic code containing the OpenDataAccessPage method in a library database, Microsoft Access looks for the form with this name, first in the library database, and then in the current database.
    - `View As AcDataAccessPageView` (optional): The view in which to open the data access page. In Access, this must be set to acDataAccessPageBrowse.
- `OpenView(ViewName As Variant, [View As AcView], [DataMode As AcOpenDataMode])`  
  The OpenView method carries out the OpenView action in Visual Basic.
    - `ViewName As Variant` (required): A string expression that's the valid name of a view in the current database. If you execute Visual Basic code containing the OpenView method in a library database, Microsoft Access looks for the view with this name first in the library database, and then in the current database.
    - `View As AcView` (optional): An AcView constant that specifies the view in which the view will open. The default value is acViewNormal.
    - `DataMode As AcOpenDataMode` (optional): An AcOpenDataMode constant that specifies the data entry mode for the view. The default value is acEdit.
- `OpenDiagram(DiagramName As Variant)`  
  The OpenDiagram method carries out the OpenDiagram action in Visual Basic.
    - `DiagramName As Variant` (required): A string expression that's the valid name of a database diagram in the current database. If you execute Visual Basic code containing the OpenDiagram method in a library database, Microsoft Access looks for the database diagram with this name first in the library database, and then in the current database.
- `OpenStoredProcedure(ProcedureName As Variant, [View As AcView], [DataMode As AcOpenDataMode])`  
  The OpenStoredProcedure method carries out the OpenStoredProcedure action in Visual Basic.
    - `ProcedureName As Variant` (required): A string expression that's the valid name of a stored procedure in the current database. If you execute Visual Basic code containing the OpenStoredProcedure method in a library database, Microsoft Access looks for the stored procedure with this name first in the library database, and then in the current database.
    - `View As AcView` (optional): An AcView constant that specifies the view in which the stored procedure will open. The default value is acViewNormal.
    - `DataMode As AcOpenDataMode` (optional): An AcOpenDataMode constant that specifies the data entry mode for the stored procedure. The default value is acEdit.
- `OpenReport(ReportName As Variant, [View As AcView], [FilterName As Variant], [WhereCondition As Variant], [WindowMode As AcWindowMode], [OpenArgs As Variant])`  
  The OpenReport method carries out the OpenReport action in Visual Basic.
    - `ReportName As Variant` (required): A string expression that's the valid name of a report in the current database. If you execute Visual Basic code containing the OpenReport method in a library database, Microsoft Access looks for the report with this name first in the library database, and then in the current database.
    - `View As AcView` (optional): An AcView constant that specifies the view in which the report will open. The default value is acViewNormal.
    - `FilterName As Variant` (optional): A string expression that's the valid name of a query in the current database.
    - `WhereCondition As Variant` (optional): A string expression that's a valid SQL WHERE clause without the word WHERE.
    - `WindowMode As AcWindowMode` (optional): An AcWindowMode constant that specifies the mode in which the form opens. The default value is acWindowNormal.
    - `OpenArgs As Variant` (optional): Sets the OpenArgs property.
- `TransferSQLDatabase(Server As Variant, Database As Variant, [UseTrustedConnection As Variant], [Login As Variant], [Password As Variant], [TransferCopyData As Variant])`  
  Transfers the entire specified Microsoft SQL Server database to another SQL Server database.
    - `Server As Variant` (required): The name of the SQL Server to which the database will be transferred.
    - `Database As Variant` (required): The name of the new database on the specified server.
    - `UseTrustedConnection As Variant` (optional): True if the current connection is using a login with system administrator privileges. If this argument is not True, you must specify a login and password in the Login and Password arguments.
    - `Login As Variant` (optional): The name of a login on the destination server with system administrator privileges. If UseTrustedConnection is True, this argument is ignored.
    - `Password As Variant` (optional): The password for the login specified in Login. If UseTrustedConnection is True, this argument is ignored.
    - `TransferCopyData As Variant` (optional): True if all data in the database is transferred to the destination database. If this argument is not True, only the database schema will be transferred.
- `CopyDatabaseFile(DatabaseFileName As Variant, [OverwriteExistingFile As Variant], [DisconnectAllUsers As Variant])`  
  Copies the database connected to the current project to a Microsoft SQL Server database file for export.
    - `DatabaseFileName As Variant` (required): The name of the file (and path) to which the current database is copied. If no path is specified, the current directory is used.
    - `OverwriteExistingFile As Variant` (optional): Determines whether Microsoft Access overwrites the file specified by DatabaseFileName. True to overwrite the existing file. If the file doesn't already exist, this argument is ignored.
    - `DisconnectAllUsers As Variant` (optional): Determines whether Access disconnects any users connected to the current database to make the copy. True to disconnect other users before copying the database file.
- `OpenFunction(FunctionName As Variant, [View As AcView], [DataMode As AcOpenDataMode])`  
  Opens a user-defined function in a Microsoft SQL Server database for viewing in Microsoft Access.
    - `FunctionName As Variant` (required): The name of the function to open.
    - `View As AcView` (optional): An AcView constant that specifies the view in which to open the function. The default value is acViewNormal.
    - `DataMode As AcOpenDataMode` (optional): An AcOpenDataMode constant that specifies the mode in which to open the function. The default value is acEdit.
- `ApplyFilter([FilterName As Variant], [WhereCondition As Variant], [ControlName As Variant])`  
  The ApplyFilter method carries out the ApplyFilter action in Visual Basic.
    - `FilterName As Variant` (optional): A string expression that is the valid name of a filter or query in the current database. When using this method to apply a server filter, the FilterName argument must be blank.
    - `WhereCondition As Variant` (optional): A string expression that is a valid SQL WHERE clause without the word WHERE.
- `OutputTo(ObjectType As AcOutputObjectType, [ObjectName As Variant], [OutputFormat As Variant], [OutputFile As Variant], [AutoStart As Variant], [TemplateFile As Variant], [Encoding As Variant], [OutputQuality As AcExportQuality])`  
  The OutputTo method carries out the OutputTo action in Visual Basic.
    - `ObjectType As AcOutputObjectType` (required): An AcOutputObjectType constant that specifies the type of object to output.
    - `ObjectName As Variant` (optional): A string expression that's the valid name of an object of the type selected by the ObjectType argument. If you want to output the active object, specify the object's type for the ObjectType argument and leave this argument blank. If you run Visual Basic code containing the OutputTo method in a library database, Microsoft Office Access searches for the object with this name first in the library database, and then in the current database.
    - `OutputFormat As Variant` (optional): An AcFormat constant that specifies the output format. If you omit this argument, Access prompts you for the output format.
    - `OutputFile As Variant` (optional): A string expression that's the full name, including the path, of the file that you want to output the object to. If you leave this argument blank, Access prompts you for an output file name.
    - `AutoStart As Variant` (optional): Use True (1) to start the appropriate Microsoft Windows-based application immediately, with the file specified by the OutputFile argument loaded. Use False (0) if you don't want to start the application. This argument is ignored for Microsoft Internet Information Server (.htx, .idc) files and Microsoft ActiveX Server (.asp) files. If you leave this argument blank, the default (False**) is assumed.
    - `TemplateFile As Variant` (optional): A string expression that's the full name, including the path, of the file that you want to use as a template for an HTML, HTX, or ASP file.
    - `Encoding As Variant` (optional): The type of character encoding format that you want used to output the text or HTML data. You can select MS-DOS, Unicode, or Unicode (UTF-8). The MS-DOS argument setting is available only for text files. If you leave this argument blank, Access outputs the data by using the Windows default encoding for text files and the default system encoding for HTML files.
    - `OutputQuality As AcExportQuality` (optional): An AcExportQuality constant that specifies the type of output device to optimize for. The default value is acExportQualityPrint.
- `TransferSharePointList(TransferType As AcSharePointListTransferType, SiteAddress As Variant, ListID As Variant, [ViewID As Variant], [TableName As Variant], [GetLookupDisplayValues As Variant])`  
  Use the TransferSharePointList method to import or link data from a SharePoint Foundation site.
    - `TransferType As AcSharePointListTransferType` (required): An AcSharePointListTransferType constant that specifies the type of transfer to make.
    - `SiteAddress As Variant` (required): The full path of the SharePoint site.
    - `ListID As Variant` (required): The name or GUID of the list to be transferred.
    - `ViewID As Variant` (optional): The GUID of the view for the list that you want to use. Leave this argument blank to transfer all rows and columns in the list.
    - `TableName As Variant` (optional): The name you want displayed for the table or linked table in Access.
    - `GetLookupDisplayValues As Variant` (optional): Specifies whether to transfer display values for Lookup fields instead of the ID used to perform the lookup.
- `CloseDatabase()`  
  Closes the current database.
- `NavigateTo([Category As Variant], [Group As Variant])`  
  Use the NavigateTo method to control the display of database objects in the navigation pane.
    - `Category As Variant` (optional): The category by which you want the navigation pane to display objects.
    - `Group As Variant` (optional): Determines which objects in the category appear in the navigation pane. If you leave this argument blank, the navigation pane will display all database objects grouped by the criteria that you specify in the Category argument. Examples of valid Group arguments for the various Category arguments are shown in the table under Remarks.
- `SearchForRecord([ObjectType As AcDataObjectType], [ObjectName As Variant], [Record As AcRecord], [WhereCondition As Variant])`  
  Use the SearchForRecord method to search for a specific record in a table, query, form, or report.
    - `ObjectType As AcDataObjectType` (optional): An AcDataObjectType constant that specifies the type of database object in which you are searching. The default value is acActiveDataObject.
    - `ObjectName As Variant` (optional): The name of the database object that contains the record to search for.
    - `Record As AcRecord` (optional): An AcRecord constant that specifies the starting point and direction of the search. The default value is acFirst.
    - `WhereCondition As Variant` (optional): A string used to locate the record. It's like the WHERE clause in an SQL statement, but without the word WHERE.
- `SetProperty([ControlName As Variant], [Property As AcProperty], [Value As Variant])`  
  The SetProperty method carries out the SetProperty action in Visual Basic.
    - `ControlName As Variant` (optional): The name of the field or control for which you want to set the property value. Leave this argument blank to set the property for the current form or report.
    - `Property As AcProperty` (optional): An AcProperty constant that specifies the property that you want to set.
    - `Value As Variant` (optional): The value to which the property is to be set. For properties whose values are either Yes or No, use 1 for Yes and 0 for No.
- `SingleStep()`  
  Pauses macro execution and opens the Macro Single Step dialog box.
- `ClearMacroError()`  
  Removes information about an error that is stored in the MacroError object.
- `SetDisplayedCategories(Show As Variant, [Category As Variant])`  
  Specifies which categories are displayed under Navigate to Category in the title bar of the navigation pane.
    - `Show As Variant` (required): Set to Yes to show the category or categories. Set to No to hide them.
    - `Category As Variant` (optional): The name of the category that you want to show or hide. Leave blank to show or hide all categories.
- `LockNavigationPane(Lock As Variant)`  
  Use the LockNavigationPane method to prevent users from deleting database objects that are displayed in the navigation pane.
    - `Lock As Variant` (required): Set to True to lock the navigation pane.
- `RunSavedImportExport(SavedImportExportName As Variant)`  
  Run a saved import or export specification.
    - `SavedImportExportName As Variant` (required): The name of a saved import or export specification to run.
- `BrowseTo(ObjectType As AcBrowseToObjectType, ObjectName As Variant, [PathtoSubformControl As Variant], [WhereCondition As Variant], [Page As Variant], [DataMode As AcFormOpenDataMode])`  
  The BrowseTo method performs the BrowseTo action in Visual Basic.
    - `ObjectType As AcBrowseToObjectType` (required): The object type to which to browse.
    - `ObjectName As Variant` (required): The object that loads inside the subform control referenced by the PathtoSubformControl argument.
    - `PathtoSubformControl As Variant` (optional): If specified, the path from the main form of the application to the target subform control that loads the object specified by the ObjectName argument.
    - `WhereCondition As Variant` (optional): If specified, replaces the Where condition of the object record source.
    - `Page As Variant` (optional): If specified, sets the page of the continuous form that will be made the current page. This argument is web only.
    - `DataMode As AcFormOpenDataMode` (optional): If specified, the data entry mode of the form.
- `SetParameter(Name As Variant, Expression As Variant)`  
  Use the SetParameter method to create a parameter for use by the BrowseTo, OpenForm, OpenQuery, OpenReport, or RunDataMacro methods.
    - `Name As Variant` (required): The name of the parameter. The name must match the name of the parameter expected by the BrowseTo, OpenForm, OpenQuery, OpenReport, or RunDataMacro method.
    - `Expression As Variant` (required): An expression that evaluates to a value to assign to the parameter.
- `RunDataMacro(MacroName As Variant)`  
  Use the RunDataMacro method to run a named data macro from Visual Basic.
    - `MacroName As Variant` (required): Name of the saved macro. The name must include the name of the table to which the data macro is attached (for example, Comments.AddComment).
- `SetOrderBy(OrderBy As Variant, [ControlName As Variant])`  
  Use the SetOrderBy method to apply a sort to the active datasheet, form, report, or table.
    - `OrderBy As Variant` (required): A string expression that includes the name of the field or fields on which to sort records and the optional ASC or DESC keywords.
    - `ControlName As Variant` (optional): If provided and the active object is a form or report, the name of the control that corresponds to the subform or subreport that will be sorted. If empty and the active object is a form or report, the parent form or report is sorted.
- `SetFilter([FilterName As Variant], [WhereCondition As Variant], [ControlName As Variant])`  
  Use the SetFilter method to apply a filter to the records in the active datasheet, form, report, or table.
    - `FilterName As Variant` (optional): If provided, the name of a query or of a filter saved as a query. This argument or the WhereCondition argument is required.
    - `WhereCondition As Variant` (optional): If provided, a SQL WHERE clause that restricts the records in the datasheet, form, report, or table.
    - `ControlName As Variant` (optional): If provided, the name of the control that corresponds to the subform or subreport to be filtered. If empty, the current object is filtered.
- `RefreshRecord()`  
  The RefreshRecord method performs the RefreshRecord macro operation from Visual Basic.
