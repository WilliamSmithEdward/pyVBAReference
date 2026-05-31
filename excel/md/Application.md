# Application

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024500-0000-0000-C000-000000000046}  

Represents the entire Microsoft Excel application.

**Remarks:** The Application object contains: - Application-wide settings and options. - Methods that return top-level objects, such as ActiveCell, ActiveSheet, and so on.

**Example:**

```vba
Application.Windows("book1.xls").Activate
```

## Properties (215)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Application  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ActiveCell As Range  (read-only)`  
  Returns a Range object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.
- `ActiveChart As Chart  (read-only)`  
  Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns Nothing.
- `ActivePrinter As String  (read/write)`  
  Returns or sets the name of the active printer. Read/write String.
- `ActiveSheet As Object  (read-only)`  
  Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns Nothing if no sheet is active.
- `ActiveWindow As Window  (read-only)`  
  Returns a Window object that represents the active Excel window (the window on top). Returns Nothing if there are no windows open. Read-only.
- `ActiveWorkbook As Workbook  (read-only)`  
  Returns a Workbook object that represents the workbook in the active window (the window on top). Returns Nothing if there are no windows open or if either the Info window or the Clipboard window is the active window. Read-only.
- `AddIns As AddIns  (read-only)`  
  Returns an AddIns collection that represents all the add-ins listed in the Add-Ins dialog box (Add-Ins command on the Developer tab). Read-only.
- `Cells As Range  (read-only)`  
  Returns a Range object that represents all the cells on the active worksheet. If the active document is not a worksheet, this property fails.
- `Charts As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the chart sheets in the active workbook.
- `Columns As Range  (read-only)`  
  Returns a Range object that represents all the columns on the active worksheet. If the active document isn't a worksheet, the Columns property fails.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars object that represents the Microsoft Excel command bars. Read-only.
- `DDEAppReturnCode As Long  (read-only)`  
  Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only Long.
- `Names As Names  (read-only)`  
  Returns a Names collection that represents all the names in the active workbook. Read-only Names object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents a cell or a range of cells.
- `Rows As Range  (read-only)`  
  Returns a Range object that represents all the rows on the active worksheet. If the active document isn't a worksheet, the Rows property fails. Read-only Range object.
- `Selection As Object  (read-only)`  
  Returns the currently selected object on the active worksheet for an Application object. Returns Nothing if no objects are selected. Use the Select method to set the selection, and use the TypeName function to discover the kind of object that is selected.
- `Sheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the sheets in the active workbook. Read-only Sheets object.
- `ThisWorkbook As Workbook  (read-only)`  
  Returns a Workbook object that represents the workbook where the current macro code is running. Read-only.
- `Windows As Windows  (read-only)`  
  Returns a Windows collection that represents all the windows in all the workbooks. Read-only Windows object.
- `Workbooks As Workbooks  (read-only)`  
  Returns a Workbooks collection that represents all the open workbooks. Read-only.
- `WorksheetFunction As WorksheetFunction  (read-only)`  
  Returns the WorksheetFunction object. Read-only.
- `Worksheets As Sheets  (read-only)`  
  For an Application object, returns a Sheets collection that represents all the worksheets in the active workbook.
- `Excel4IntlMacroSheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.
- `Excel4MacroSheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.
- `AlertBeforeOverwriting As Boolean  (read/write)`  
  True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation. Read/write Boolean.
- `AltStartupPath As String  (read/write)`  
  Returns or sets the name of the alternate startup folder. Read/write String.
- `AskToUpdateLinks As Boolean  (read/write)`  
  True if Microsoft Excel asks the user to update links when opening files with links. False if links are automatically updated with no dialog box. Read/write Boolean.
- `EnableAnimations As Boolean  (read/write)`
- `AutoCorrect As AutoCorrect  (read-only)`  
  Returns an AutoCorrect object that represents the Microsoft Excel AutoCorrect attributes. Read-only.
- `Build As Long  (read-only)`  
  Returns the Microsoft Excel build number. Read-only Long.
- `CalculateBeforeSave As Boolean  (read/write)`  
  True if workbooks are calculated before they're saved to disk (if the Calculation property is set to xlManual). This property is preserved even if you change the Calculation property. Read/write Boolean.
- `Calculation As XlCalculation  (read/write)`  
  Returns or sets an XlCalculation value that represents the calculation mode.
- `Caller As Variant  (read-only)`  
  Returns information about how Visual Basic was called (for more information, see the Remarks section).
- `CanPlaySounds As Boolean  (read-only)`
- `CanRecordSounds As Boolean  (read-only)`
- `Caption As String  (read/write)`  
  Returns or sets a String value that represents the name that appears in the title bar of the main Microsoft Excel window.
- `CellDragAndDrop As Boolean  (read/write)`  
  True if dragging and dropping cells is enabled. Read/write Boolean.
- `ClipboardFormats As Variant  (read-only)`  
  Returns the formats that are currently on the Clipboard, as an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element in the array with the appropriate constant listed in the Remarks section. Read-only Variant.
- `DisplayClipboardWindow As Boolean  (read/write)`  
  Returns True if the Microsoft Office Clipboard can be displayed. Read/write Boolean.
- `CommandUnderlines As XlCommandUnderlines  (read/write)`  
  Returns or sets the state of the command underlines in Microsoft Excel for the Macintosh. Can be one of the constants of XlCommandUnderlines. Read/write Long.
- `ConstrainNumeric As Boolean  (read/write)`  
  True if handwriting recognition is limited to numbers and punctuation only. Read/write Boolean.
- `CopyObjectsWithCells As Boolean  (read/write)`  
  True if objects are cut, copied, extracted, and sorted with cells. Read/write Boolean.
- `Cursor As XlMousePointer  (read/write)`  
  Returns or sets the appearance of the mouse pointer in Microsoft Excel. Read/write XlMousePointer.
- `CustomListCount As Long  (read-only)`  
  Returns the number of defined custom lists (including built-in lists). Read-only Long.
- `CutCopyMode As XlCutCopyMode  (read/write)`  
  Returns or sets the status of Cut or Copy mode. Can be True, False, or an XLCutCopyMode constant, as shown in the following tables. Read/write Long.
- `DataEntryMode As Long  (read/write)`  
  Returns or sets Data Entry mode, as shown in the following table. When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range. Read/write Long.
- `_Default As String  (read-only)`
- `DefaultFilePath As String  (read/write)`  
  Returns or sets the default path that Microsoft Excel uses when it opens files. Read/write String.
- `Dialogs As Dialogs  (read-only)`  
  Returns a Dialogs collection that represents all built-in dialog boxes. Read-only.
- `DisplayAlerts As Boolean  (read/write)`  
  True if Microsoft Excel displays certain alerts and messages while a macro is running. Read/write Boolean.
- `DisplayFormulaBar As Boolean  (read/write)`  
  True if the formula bar is displayed. Read/write Boolean.
- `DisplayFullScreen As Boolean  (read/write)`  
  True if Microsoft Excel is in full-screen mode. Read/write Boolean.
- `DisplayNoteIndicator As Boolean  (read/write)`  
  True if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners). Read/write Boolean.
- `DisplayCommentIndicator As XlCommentDisplayMode  (read/write)`  
  Returns or sets the way cells display comments and indicators. Can be one of the XlCommentDisplayMode constants.
- `DisplayExcel4Menus As Boolean  (read/write)`  
  True if Microsoft Excel displays version 4.0 menu bars. Read/write Boolean.
- `DisplayRecentFiles As Boolean  (read/write)`  
  True if the list of recently used files is displayed in the UI. Read/write Boolean.
- `DisplayScrollBars As Boolean  (read/write)`  
  True if scroll bars are visible for all workbooks. Read/write Boolean.
- `DisplayStatusBar As Boolean  (read/write)`  
  True if the status bar is displayed. Read/write Boolean.
- `EditDirectlyInCell As Boolean  (read/write)`  
  True if Microsoft Excel allows editing in cells. Read/write Boolean.
- `EnableAutoComplete As Boolean  (read/write)`  
  True if the AutoComplete feature is enabled. Read/write Boolean.
- `EnableCancelKey As XlEnableCancelKey  (read/write)`  
  Controls how Microsoft Excel handles Ctrl+Break (or Esc or Command+Period) user interruptions to the running procedure. Read/write XlEnableCancelKey.
- `EnableSound As Boolean  (read/write)`  
  True if sound is enabled for Microsoft Office. Read/write Boolean.
- `FileConverters As Variant  (read-only)`  
  Returns information about installed file converters. Returns null if there are no converters installed. Read-only Variant.
- `FixedDecimal As Boolean  (read/write)`  
  All data entered after this property is set to True will be formatted with the number of fixed decimal places set by the FixedDecimalPlaces property. Read/write Boolean.
- `FixedDecimalPlaces As Long  (read/write)`  
  Returns or sets the number of fixed decimal places used when the FixedDecimal property is set to True. Read/write Long.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the main application window.
- `IgnoreRemoteRequests As Boolean  (read/write)`  
  True if remote DDE requests are ignored. Read/write Boolean.
- `Interactive As Boolean  (read/write)`  
  True if Microsoft Excel is in interactive mode; this property is usually True. If you set this property to False, Excel blocks all input from the keyboard and mouse (except input to dialog boxes that are displayed by your code). Read/write Boolean.
- `International As Variant  (read-only)`  
  Returns information about the current country/region and international settings. Read-only Variant.
- `Iteration As Boolean  (read/write)`  
  True if Microsoft Excel uses iteration to resolve circular references. Read/write Boolean.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.
- `LibraryPath As String  (read-only)`  
  Returns the path to the Library folder, but without the final separator. Read-only String.
- `MailSession As Variant  (read-only)`  
  Returns the MAPI mail session number as a hexadecimal string (if there's an active session), or returns null if there's no session. Read-only Variant.
- `MailSystem As XlMailSystem  (read-only)`  
  Returns the mail system that's installed on the host machine. Read-only XlMailSystem.
- `MathCoprocessorAvailable As Boolean  (read-only)`  
  True if a math coprocessor is available. Read-only Boolean.
- `MaxChange As Double  (read/write)`  
  Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. Read/write Double.
- `MaxIterations As Long  (read/write)`  
  Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference. Read/write Long.
- `MouseAvailable As Boolean  (read-only)`  
  True if a mouse is available. Read-only Boolean.
- `MoveAfterReturn As Boolean  (read/write)`  
  True if the active cell is moved as soon as the Enter (Return) key is pressed. Read/write Boolean.
- `MoveAfterReturnDirection As XlDirection  (read/write)`  
  Returns or sets the direction in which the active cell is moved when the user presses Enter. Read/write XlDirection.
- `RecentFiles As RecentFiles  (read-only)`  
  Returns a RecentFiles collection that represents the list of recently used files.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `NetworkTemplatesPath As String  (read-only)`  
  Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. Read-only String.
- `ODBCErrors As ODBCErrors  (read-only)`  
  Returns an ODBCErrors collection that contains all the ODBC errors generated by the most recent query table or PivotTable report operation. Read-only.
- `ODBCTimeout As Long  (read/write)`  
  Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. Read/write Long.
- `OnWindow As String  (read/write)`  
  Returns or sets the name of the procedure that's run whenever you activate a window. Read/write String.
- `OperatingSystem As String  (read-only)`  
  Returns the name and version number of the current operating system. Read-only String.
- `OrganizationName As String  (read-only)`  
  Returns the registered organization name. Read-only String.
- `Path As String  (read-only)`  
  Returns a String value that represents the complete path to the application, excluding the final separator and name of the application.
- `PathSeparator As String  (read-only)`  
  Returns the path separator character (\). Read-only String.
- `PreviousSelections As Variant  (read-only)`  
  Returns an array of the last four ranges or names selected. Each element in the array is a Range object. Read-only Variant.
- `PivotTableSelection As Boolean  (read/write)`  
  True if PivotTable reports use structured selection. Read/write Boolean.
- `PromptForSummaryInfo As Boolean  (read/write)`  
  True if Microsoft Excel asks for summary information when files are first saved. Read/write Boolean.
- `RecordRelative As Boolean  (read-only)`  
  True if macros are recorded by using relative references; False if recording is absolute. Read-only Boolean.
- `ReferenceStyle As XlReferenceStyle  (read/write)`  
  Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. Read/write XlReferenceStyle.
- `RegisteredFunctions As Variant  (read-only)`  
  Returns information about functions in either dynamic-link libraries (DLLs) or code resources that were registered with the REGISTER or REGISTER.ID macro functions. Read-only Variant.
- `RollZoom As Boolean  (read/write)`  
  True if the IntelliMouse zooms instead of scrolling. Read/write Boolean.
- `ScreenUpdating As Boolean  (read/write)`  
  True if screen updating is turned on. Read/write Boolean.
- `SheetsInNewWorkbook As Long  (read/write)`  
  Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks. Read/write Long.
- `ShowChartTipNames As Boolean  (read/write)`  
  True if charts show chart tip names. The default value is True. Read/write Boolean.
- `ShowChartTipValues As Boolean  (read/write)`  
  True if charts show chart tip values. The default value is True. Read/write Boolean.
- `StandardFont As String  (read/write)`  
  Returns or sets the name of the standard font. Read/write String.
- `StandardFontSize As Double  (read/write)`  
  Returns or sets the standard font size, in points. Read/write Long.
- `StartupPath As String  (read-only)`  
  Returns the complete path of the startup folder, excluding the final separator. Read-only String.
- `StatusBar As Variant  (read/write)`  
  Returns or sets the text in the status bar. Read/write String.
- `TemplatesPath As String  (read-only)`  
  Returns the local path where templates are stored. Read-only String.
- `ShowToolTips As Boolean  (read/write)`  
  True if ToolTips are turned on. Read/write Boolean.
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.
- `DefaultSaveFormat As XlFileFormat  (read/write)`  
  Returns or sets the default format for saving files. For a list of valid constants, see the FileFormat property. Read/write Long.
- `TransitionMenuKey As String  (read/write)`  
  Returns or sets the Microsoft Excel menu or help key, which is usually /. Read/write String.
- `TransitionMenuKeyAction As Long  (read/write)`  
  Returns or sets the action taken when the Microsoft Excel menu key is pressed. Can be either xlExcelMenus or xlLotusHelp (see the Excel constants enumeration). Read/write Long.
- `TransitionNavigKeys As Boolean  (read/write)`  
  True if transition navigation keys are active. Read/write Boolean.
- `UsableHeight As Double  (read-only)`  
  Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.
- `UsableWidth As Double  (read-only)`  
  Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only Double.
- `UserControl As Boolean  (read/write)`  
  True if the application is visible or if it was created or started by the user. False if you created or started the application programmatically by using the CreateObject or GetObject functions, and the application is hidden. Read/write Boolean.
- `UserName As String  (read/write)`  
  Returns or sets the name of the current user. Read/write String.
- `Value As String  (read-only)`  
  Returns a String value that represents the name of the application.
- `VBE As VBE  (read-only)`  
  Returns a VBE object that represents the Visual Basic Editor. Read-only.
- `Version As String  (read-only)`  
  Returns a String value that represents the Microsoft Excel version number.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the application window to its right edge.
- `WindowsForPens As Boolean  (read-only)`  
  True if the computer is running under Microsoft Windows for Pen Computing. Read-only Boolean.
- `WindowState As XlWindowState  (read/write)`  
  Returns or sets the state of the window. Read/write XlWindowState.
- `DefaultSheetDirection As Long  (read/write)`  
  Returns or sets the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following XlReadingOrder constants: xlRTL (right to left) or xlLTR (left to right). Read/write Long.
- `CursorMovement As Long  (read/write)`  
  Returns or sets a value that indicates whether a visual cursor or a logical cursor is used. Can be one of the following constants: xlVisualCursor or xlLogicalCursor. Read/write Long.
- `ControlCharacters As Boolean  (read/write)`  
  True if Microsoft Excel displays control characters for right-to-left languages. Read/write Boolean.
- `EnableEvents As Boolean  (read/write)`  
  True if events are enabled for the specified object. Read/write Boolean.
- `ExtendList As Boolean  (read/write)`  
  True if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list. Read/write Boolean.
- `OLEDBErrors As OLEDBErrors  (read-only)`  
  Returns the OLEDBErrors collection, which represents the error information returned by the most recent OLE DB query. Read-only.
- `COMAddIns As COMAddIns  (read-only)`  
  Returns the COMAddIns collection for Microsoft Excel, which represents the currently installed COM add-ins. Read-only.
- `DefaultWebOptions As DefaultWebOptions  (read-only)`  
  Returns the DefaultWebOptions object that contains global application-level attributes used by Microsoft Excel whenever you save a document as a webpage or open a webpage. Read-only.
- `ProductCode As String  (read-only)`  
  Returns the globally unique identifier (GUID) for Microsoft Excel. Read-only String.
- `UserLibraryPath As String  (read-only)`  
  Returns the path to the location on the user's computer where the COM add-ins are installed. Read-only String.
- `AutoPercentEntry As Boolean  (read/write)`  
  True if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered. Read/write Boolean.
- `LanguageSettings As LanguageSettings  (read-only)`  
  Returns the LanguageSettings object, which contains information about the language settings in Microsoft Excel. Read-only.
- `CalculationVersion As Long  (read-only)`  
  Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits (on the left) are the major version of Microsoft Excel. Read-only Long.
- `FeatureInstall As MsoFeatureInstall  (read/write)`  
  Returns or sets a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the MsoFeatureInstall constants listed in the following table. Read/write MsoFeatureInstall.
- `Ready As Boolean  (read-only)`  
  Returns True when the Microsoft Excel application is ready; False when the Excel application is not ready. Read-only Boolean.
- `FindFormat As CellFormat  (read/write)`  
  Sets or returns the search criteria for the type of cell formats to find.
- `ReplaceFormat As CellFormat  (read/write)`  
  Sets the replacement criteria to use in replacing cell formats. The replacement criteria is then used in a subsequent call to the Replace method of the Range object.
- `UsedObjects As UsedObjects  (read-only)`  
  Returns a UsedObjects object representing objects allocated in a workbook. Read-only.
- `CalculationState As XlCalculationState  (read-only)`  
  Returns an XlCalculationState constant that indicates the calculation state of the application, for any calculations that are being performed in Microsoft Excel. Read-only.
- `CalculationInterruptKey As XlCalculationInterruptKey  (read/write)`  
  Sets or returns an XlCalculationInterruptKey constant that specifies the key that can interrupt Microsoft Excel when performing calculations. Read/write.
- `Watches As Watches  (read-only)`  
  Returns a Watches object representing a range that is tracked when the worksheet is recalculated.
- `DisplayFunctionToolTips As Boolean  (read/write)`  
  True if function ToolTips can be displayed. Read/write Boolean.
- `AutomationSecurity As MsoAutomationSecurity  (read/write)`  
  Returns or sets an MsoAutomationSecurity constant that represents the security mode that Microsoft Excel uses when programmatically opening files. Read/write.
- `FileDialog As FileDialog  (read-only)`  
  Returns a FileDialog object representing an instance of the file dialog.
- `DisplayPasteOptions As Boolean  (read/write)`  
  True if the Paste Options button can be displayed. Read/write Boolean.
- `DisplayInsertOptions As Boolean  (read/write)`  
  True if the Insert Options button should be displayed. Read/write Boolean.
- `GenerateGetPivotData As Boolean  (read/write)`  
  Returns True when Microsoft Excel can get PivotTable report data. Read/write Boolean.
- `AutoRecover As AutoRecover  (read-only)`  
  Returns an AutoRecover object, which backs up all file formats on a timed interval.
- `Hwnd As Long  (read-only)`  
  Returns a Long indicating the top-level window handle of the Microsoft Excel window. Read-only.
- `Hinstance As Long  (read-only)`  
  Returns a handle to the instance of Excel represented by the Application object. Read-only Long.
- `ErrorCheckingOptions As ErrorCheckingOptions  (read-only)`  
  Returns an ErrorCheckingOptions object, which represents the error checking options for an application.
- `AutoFormatAsYouTypeReplaceHyperlinks As Boolean  (read/write)`  
  True (default) if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type. Read/write Boolean.
- `NewWorkbook As NewFile  (read-only)`  
  Returns a NewFile object.
- `SpellingOptions As SpellingOptions  (read-only)`  
  Returns a SpellingOptions object that represents the spelling options of the application.
- `Speech As Speech  (read-only)`  
  Returns a Speech object.
- `MapPaperSize As Boolean  (read/write)`  
  True if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region. Read/write Boolean.
- `ShowStartupDialog As Boolean  (read/write)`  
  Returns True (default is False) when the New Workbook task pane appears for a Microsoft Excel application. Read/write Boolean.
- `DecimalSeparator As String  (read/write)`  
  Sets or returns the character used for the decimal separator as a String. Read/write.
- `ThousandsSeparator As String  (read/write)`  
  Sets or returns the character used for the thousands separator as a String. Read/write.
- `UseSystemSeparators As Boolean  (read/write)`  
  True (default) if the system separators of Microsoft Excel are enabled. Read/write Boolean.
- `ThisCell As Range  (read-only)`  
  Returns the cell in which the user-defined function is being called from as a Range object.
- `RTD As RTD  (read-only)`  
  Returns an RTD object.
- `DisplayDocumentActionTaskPane As Boolean  (read/write)`  
  Set to True to display the Document Actions task pane; set to False to hide the Document Actions task pane. Read/write Boolean.
- `ArbitraryXMLSupportAvailable As Boolean  (read-only)`  
  Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available. Read-only.
- `MeasurementUnit As Long  (read/write)`  
  Specifies the measurement unit used in the application. Read/write XlMeasurementUnits.
- `ShowSelectionFloaties As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Mini toolbars displays when a user selects text. False if Mini toolbars are displayed. Read/write Boolean.
- `ShowMenuFloaties As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to display Mini toolbars when the user right-clicks in the workbook window. False if Mini toolbars are displayed. Read/write Boolean.
- `ShowDevTools As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the Developer tab is displayed in the ribbon. Read/write Boolean.
- `EnableLivePreview As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to True shows a preview of your workbook before applying the command. Read/write Boolean.
- `DisplayDocumentInformationPanel As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the document properties panel is displayed. Read/write Boolean.
- `AlwaysUseClearType As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to use ClearType to display fonts in the menu, ribbon, and dialog box text. Read/write Boolean.
- `WarnOnFunctionNameConflict As Boolean  (read/write)`  
  The WarnOnFunctionNameConflict property, when set to True, raises an alert if a developer tries to create a new function by using an existing function name. Read/write Boolean.
- `FormulaBarHeight As Long  (read/write)`  
  Allows the user to specify the height of the formula bar in lines. Read/write Long.
- `DisplayFormulaAutoComplete As Boolean  (read/write)`  
  Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Read/write Boolean.
- `GenerateTableRefs As XlGenerateTableRefs  (read/write)`  
  The GenerateTableRefs property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Read/write.
- `Assistance As IAssistance  (read-only)`  
  Returns an IAssistance object for Microsoft Excel that represents the Microsoft Office Help Viewer. Read-only.
- `EnableLargeOperationAlert As Boolean  (read/write)`  
  Sets or returns a Boolean that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office Center UI. Read/write Boolean.
- `LargeOperationCellThousandCount As Long  (read/write)`  
  Returns or sets the maximum number of cells needed in an operation beyond which an alert is triggered. Read/write Long.
- `DeferAsyncQueries As Boolean  (read/write)`  
  Gets or sets whether asynchronous queries to OLAP data sources are executed when a worksheet is calculated by VBA code. Read/write Boolean.
- `MultiThreadedCalculation As MultiThreadedCalculation  (read-only)`  
  Returns a MultiThreadedCalculation object that controls the multi-threaded recalculation settings. Read-only.
- `ActiveEncryptionSession As Long  (read-only)`  
  Returns a Long that represents the encryption session associated with the active document. Read-only.
- `HighQualityModeForGraphics As Boolean  (read/write)`  
  Returns or sets whether Excel uses high quality mode to print graphics. Read/write.
- `FileExportConverters As FileExportConverters  (read-only)`  
  Returns a FileExportConverters collection that represents all the file converters for saving files available to Microsoft Excel. Read-only.
- `SmartArtLayouts As SmartArtLayouts  (read-only)`  
  Returns the set of SmartArtLayouts that are currently loaded in the application. Read-only.
- `SmartArtQuickStyles As SmartArtQuickStyles  (read-only)`  
  Returns the set of SmartArtQuickStyles that are currently loaded in the application. Read-only.
- `SmartArtColors As SmartArtColors  (read-only)`  
  Returns the set of SmartArtColors styles that are currently loaded in the application. Read-only.
- `AddIns2 As AddIns2  (read-only)`  
  Returns an AddIns2 collection that represents all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed. Read-only.
- `PrintCommunication As Boolean  (read/write)`  
  Specifies whether communication with the printer is turned on. Read/write Boolean.
- `UseClusterConnector As Boolean  (read/write)`  
  Returns or sets whether Excel allows user-defined functions in XLL add-ins to be run on a compute cluster. Read/write.
- `ClusterConnector As String  (read/write)`  
  Returns or sets the name of the High Performance Computing (HPC) Cluster Connector that is used to run user-defined functions in XLL add-ins. Read/write.
- `ProtectedViewWindows As ProtectedViewWindows  (read-only)`  
  Returns a ProtectedViewWindows collection that represents all the Protected View windows that are open in the application. Read-only.
- `ActiveProtectedViewWindow As ProtectedViewWindow  (read-only)`  
  Returns a ProtectedViewWindow object that represents the active Protected View window (the window on top). Returns Nothing if there are no Protected View windows open. Read-only.
- `IsSandboxed As Boolean  (read-only)`  
  Returns True if the specified workbook is open in a Protected View window. Read-only.
- `HinstancePtr As Variant  (read-only)`  
  Returns a handle to the instance of Excel represented by the specified Application object. Read-only Variant.
- `FileValidation As MsoFileValidationMode  (read/write)`  
  Returns or sets how Excel will validate files before opening them. Read/write.
- `FileValidationPivot As XlFileValidationPivotMode  (read/write)`  
  Returns or sets how Excel will validate the contents of the data caches for PivotTable reports. Read/write.
- `ShowQuickAnalysis As Boolean  (read/write)`  
  Controls whether the Quick Analysis contextual user interface is displayed on selection. True means that the Quick Analysis button will show.
- `QuickAnalysis As QuickAnalysis  (read-only)`  
  Returns a QuickAnalysis object that represents the Quick Analysis options of the application.
- `FlashFill As Boolean  (read/write)`  
  True indicates that the Excel Flash Fill feature has been enabled and active. Read/write Boolean.
- `EnableMacroAnimations As Boolean  (read/write)`  
  Controls whether macro animations are enabled. True if user interface animations or chart animations are enabled. Is set to False (no animation) by default. If it is set to True during the running of a macro, it will enable animation, and then will reset to False after the macro runs. Read/write Boolean.
- `ChartDataPointTrack As Boolean  (read/write)`  
  True causes all charts in newly created documents to use the cell reference tracking behavior. Boolean.
- `FlashFillMode As Boolean  (read/write)`  
  True if the Flash Fill feature is enabled. Read/write Boolean.
- `MergeInstances As Boolean  (read/write)`  
  True to merge multiple instances of the application into a single instance. Read/write Boolean.
- `EnableCheckFileExtensions As Boolean  (read/write)`  
  True to enable the Tell me if Microsoft Excel isn't the default program for viewing and editing spreadsheets dialog box. Read/write Boolean.
- `DefaultPivotTableLayoutOptions As DefaultPivotTableLayoutOptions  (read-only)`
- `ShowConvertToDataType As Boolean  (read/write)`
- `SensitivityLabelPolicy As SensitivityLabelPolicy  (read-only)`  
  Returns the sensitivity label policy for the Application object.
- `FormatStaleValues As Boolean  (read/write)`
- `MaxSupportedCompatibilityVersion As Long  (read-only)`

## Methods (52)

- `Calculate()`  
  Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.
- `DDEExecute(Channel As Long, String As String)`  
  Runs a command or performs some other action or actions in another application by way of the specified DDE channel.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `String As String` (required): The message defined in the receiving application.
- `DDEInitiate(App As String, Topic As String) As Long`  
  Opens a DDE channel to an application.
    - `App As String` (required): The application name.
    - `Topic As String` (required): Describes something in the application to which you are opening a channel, usually a document of that application.
- `DDEPoke(Channel As Long, Item As Variant, Data As Variant)`  
  Sends data to an application.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `Item As Variant` (required): The item to which the data is to be sent.
    - `Data As Variant` (required): The data to be sent to the application.
- `DDERequest(Channel As Long, Item As String) As Variant`  
  Requests information from the specified application. This method always returns an array.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `Item As String` (required): The item to be requested.
- `DDETerminate(Channel As Long)`  
  Closes a channel to another application.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
- `Evaluate(Name As Variant) As Variant`  
  Converts a Microsoft Excel name to an object or a value.
    - `Name As Variant` (required): A formula or the name of the object, using the naming convention of Microsoft Excel. The length of the name must be less than or equal to 255 characters.
- `_Evaluate(Name As Variant) As Variant`
- `ExecuteExcel4Macro(String As String) As Variant`  
  Runs a Microsoft Excel 4.0 macro function and then returns the result of the function. The return type depends on the function.
    - `String As String` (required): A Microsoft Excel 4.0 macro language function without the equal sign. All references must be given as R1C1 strings. If _String_ contains embedded double quotation marks, you must double them. For example, to run the macro function =MID("sometext",1,4), _String_ would have to be "MID(""sometext"",1,4)".
- `Intersect(Arg1 As Range, Arg2 As Range, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Range`  
  Returns a Range object that represents the rectangular intersection of two or more ranges. If one or more ranges from a different worksheet are specified, an error is returned.
    - `Arg1 As Range` (required): The intersecting ranges. At least two Range objects must be specified.
    - `Arg2 As Range` (required): The intersecting ranges. At least two Range objects must be specified.
- `Run([Macro As Variant], [Arg1 As Variant], [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Variant`  
  Runs a macro or calls a function. This can be used to run a macro written in Visual Basic or the Microsoft Excel macro language, or to run a function in a DLL or XLL.
    - `Macro As Variant` (optional): The macro to run. This can be either a string with the macro name, a Range object indicating where the function is, or a register ID for a registered DLL (XLL) function. If a string is used, the string will be evaluated in the context of the active sheet.
- `_Run2([Macro As Variant], [Arg1 As Variant], [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Variant`
- `SendKeys(Keys As Variant, [Wait As Variant])`  
  Sends keystrokes to the active application.
    - `Keys As Variant` (required): The key or key combination that you want to send to the application, as text.
    - `Wait As Variant` (optional): True to have Microsoft Excel wait for the keys to be processed before returning control to the macro. False (or omitted) to continue running the macro without waiting for the keys to be processed.
- `Union(Arg1 As Range, Arg2 As Range, [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Range`  
  Returns the union of two or more ranges.
    - `Arg1 As Range` (required): At least two Range objects must be specified.
    - `Arg2 As Range` (required): At least two Range objects must be specified.
- `ActivateMicrosoftApp(Index As XlMSApplication)`  
  Activates a Microsoft application. If the application is already running, this method activates the running application. If the application isn't running, this method starts a new instance of the application.
    - `Index As XlMSApplication` (required): Specifies the Microsoft application to activate.
- `AddCustomList(ListArray As Variant, [ByRow As Variant])`  
  Adds a custom list for custom autofill and/or custom sort.
    - `ListArray As Variant` (required): Specifies the source data, as either an array of strings or a Range object.
    - `ByRow As Variant` (optional): Only used if _ListArray_ is a Range object. True to create a custom list from each row in the range. False to create a custom list from each column in the range. If this argument is omitted and there are more rows than columns (or an equal number of rows and columns) in the range, Microsoft Excel creates a custom list from each column in the range. If this argument is omitted and there are more columns than rows in the range, Excel creates a custom list from each row in the range.
- `CentimetersToPoints(Centimeters As Double) As Double`  
  Converts a measurement from centimeters to points (one point equals 0.035 centimeters).
    - `Centimeters As Double` (required): Specifies the centimeter value to be converted to points.
- `CheckSpelling(Word As String, [CustomDictionary As Variant], [IgnoreUppercase As Variant]) As Boolean`  
  Checks the spelling of a single word.
    - `Word As String` (required): (used only with the Application object). The word that you want to check.
    - `CustomDictionary As Variant` (optional): A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.
    - `IgnoreUppercase As Variant` (optional): True to have Microsoft Excel ignore words that are all uppercase. False to have Microsoft Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.
- `ConvertFormula(Formula As Variant, FromReferenceStyle As XlReferenceStyle, [ToReferenceStyle As Variant], [ToAbsolute As Variant], [RelativeTo As Variant]) As Variant`  
  Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both. Variant.
    - `Formula As Variant` (required): A string that contains the formula that you want to convert. This must be a valid formula, and it must begin with an equal sign.
    - `FromReferenceStyle As XlReferenceStyle` (required): The reference style of the formula.
    - `ToReferenceStyle As Variant` (optional): A constant of XlReferenceStyle specifying the reference style that you want returned. If this argument is omitted, the reference style isn't changed; the formula stays in the style specified by _FromReferenceStyle_.
    - `ToAbsolute As Variant` (optional): A constant of XlReferenceType that specifies the converted reference type. If this argument is omitted, the reference type isn't changed.
    - `RelativeTo As Variant` (optional): A Range object that contains one cell. Relative references relate to this cell.
- `DeleteCustomList(ListNum As Long)`  
  Deletes a custom list.
    - `ListNum As Long` (required): The custom list number. This number must be greater than or equal to 5 (Microsoft Excel has four built-in custom lists that cannot be deleted).
- `DoubleClick()`  
  Equivalent to double-clicking the active cell.
- `GetCustomListContents(ListNum As Long) As Variant`  
  Returns a custom list (an array of strings).
    - `ListNum As Long` (required): The list number.
- `GetCustomListNum(ListArray As Variant) As Long`  
  Returns the custom list number for an array of strings. Use this method to match both built-in lists and custom-defined lists.
    - `ListArray As Variant` (required): An array of strings.
- `GetOpenFilename([FileFilter As Variant], [FilterIndex As Variant], [Title As Variant], [ButtonText As Variant], [MultiSelect As Variant]) As Variant`  
  Displays the standard Open dialog box and gets a file name from the user without actually opening any files.
    - `FileFilter As Variant` (optional): A string specifying file filtering criteria.
    - `FilterIndex As Variant` (optional): Specifies the index numbers of the default file filtering criteria, from 1 to the number of filters specified in _FileFilter_. If this argument is omitted or greater than the number of filters present, the first file filter is used.
    - `Title As Variant` (optional): Specifies the title of the dialog box. If this argument is omitted, the title is "Open."
    - `ButtonText As Variant` (optional): Macintosh only.
    - `MultiSelect As Variant` (optional): True to allow multiple file names to be selected. False to allow only one file name to be selected. The default value is False.
- `GetSaveAsFilename([InitialFilename As Variant], [FileFilter As Variant], [FilterIndex As Variant], [Title As Variant], [ButtonText As Variant]) As Variant`  
  Displays the standard Save As dialog box and gets a file name from the user without actually saving any files.
    - `InitialFilename As Variant` (optional): Specifies the suggested file name. If this argument is omitted, Microsoft Excel uses the active workbook's name.
    - `FileFilter As Variant` (optional): A string specifying file filtering criteria. Max length is 255 characters, otherwise the method returns Error 2015.
    - `FilterIndex As Variant` (optional): Specifies the index number of the default file filtering criteria, from 1 to the number of filters specified in _FileFilter_. If this argument is omitted or greater than the number of filters present, the first file filter is used.
    - `Title As Variant` (optional): Specifies the title of the dialog box. If this argument is omitted, the default title is used.
    - `ButtonText As Variant` (optional): Macintosh only.
- `Goto([Reference As Variant], [Scroll As Variant])`  
  Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it's not already active.
    - `Reference As Variant` (optional): The destination. Can be a Range object, a string that contains a cell reference in R1C1-style notation, or a string that contains a Visual Basic procedure name. If this argument is omitted, the destination is the last range you used the Goto method to select.
    - `Scroll As Variant` (optional): True to scroll through the window so that the upper-left corner of the range appears in the upper-left corner of the window. False to not scroll through the window. The default is False.
- `Help([HelpFile As Variant], [HelpContextID As Variant])`  
  Displays a Help topic.
    - `HelpFile As Variant` (optional): The name of the online Help file that you want to display. If this argument isn't specified, Microsoft Excel Help is used.
    - `HelpContextID As Variant` (optional): Specifies the context ID number for the Help topic. If this argument isn't specified, the Help Topics dialog box is displayed.
- `InchesToPoints(Inches As Double) As Double`  
  Converts a measurement from inches to points.
    - `Inches As Double` (required): Specifies the inch value to be converted to points.
- `InputBox(Prompt As String, [Title As Variant], [Default As Variant], [Left As Variant], [Top As Variant], [HelpFile As Variant], [HelpContextID As Variant], [Type As Variant]) As Variant`  
  Displays a dialog box for user input. Returns the information entered in the dialog box.
    - `Prompt As String` (required): The message to be displayed in the dialog box. This can be a string, a number, a date, or a Boolean value (Microsoft Excel automatically coerces the value to a String before it is displayed). Maximum length is 255 characters, otherwise there is no prompt, and Application's method immediately returns Error 2015.
    - `Title As Variant` (optional): The title for the input box. If this argument is omitted, the default title is Input.
    - `Default As Variant` (optional): Specifies a value that will appear in the text box when the dialog box is initially displayed. If this argument is omitted, the text box is left empty. This value can be a Range object.
    - `Left As Variant` (optional): Specifies an x position for the dialog box in relation to the upper-left corner of the screen, in points.
    - `Top As Variant` (optional): Specifies a y position for the dialog box in relation to the upper-left corner of the screen, in points.
    - `HelpFile As Variant` (optional): The name of the Help file for this input box. If the _HelpFile_ and _HelpContextID_ arguments are present, a Help button will appear in the dialog box.
    - `HelpContextID As Variant` (optional): The context ID number of the Help topic in _HelpFile_.
    - `Type As Variant` (optional): Specifies the return data type. If this argument is omitted, the dialog box returns text.
- `MailLogoff()`  
  Closes a MAPI mail session established by Microsoft Excel.
- `MailLogon([Name As Variant], [Password As Variant], [DownloadNewMail As Variant])`  
  Logs on to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail isn't already running, you must use this method to establish a mail session before mail or document routing functions can be used.
    - `Name As Variant` (optional): The mail account name or Microsoft Exchange profile name. If this argument is omitted, the default mail account name is used.
    - `Password As Variant` (optional): The mail account password. This argument is ignored in Microsoft Exchange.
    - `DownloadNewMail As Variant` (optional): True to download new mail immediately.
- `NextLetter() As Workbook`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `OnKey(Key As String, [Procedure As Variant])`  
  Runs a specified procedure when a particular key or key combination is pressed.
    - `Key As String` (required): A string indicating the key to be pressed.
    - `Procedure As Variant` (optional): A string indicating the name of the procedure to be run. If _Procedure_ is "" (empty text), nothing happens when _Key_ is pressed. This form of OnKey changes the normal result of keystrokes in Microsoft Excel. If _Procedure_ is omitted, _Key_ reverts to its normal result in Microsoft Excel, and any special key assignments made with previous OnKey methods are cleared.
- `OnRepeat(Text As String, Procedure As String)`  
  Sets the Repeat item and the name of the procedure that will run if you choose the Repeat command after running the procedure that sets this property.
    - `Text As String` (required): The text that appears with the Repeat command.
    - `Procedure As String` (required): The name of the procedure that will be run when you choose the Repeat command.
- `OnTime(EarliestTime As Variant, Procedure As String, [LatestTime As Variant], [Schedule As Variant])`  
  Schedules a procedure to be run at a specified time in the future (either at a specific time of day or after a specific amount of time has passed).
    - `EarliestTime As Variant` (required): The time when you want this procedure to be run.
    - `Procedure As String` (required): The name of the procedure to be run.
    - `LatestTime As Variant` (optional): The latest time at which the procedure can be run. For example, if _LatestTime_ is set to _EarliestTime_ + 30 and Microsoft Excel is not in Ready, Copy, Cut, or Find mode at _EarliestTime_ because another procedure is running, Excel will wait 30 seconds for the first procedure to complete. If Excel is not in Ready mode within 30 seconds, the procedure won't be run. If this argument is omitted, Excel will wait until the procedure can be run.
    - `Schedule As Variant` (optional): True to schedule a new OnTime procedure. False to clear a previously set procedure. The default value is True.
- `OnUndo(Text As String, Procedure As String)`  
  Sets the text of the Undo command and the name of the procedure that's run if you choose the Undo command after running the procedure that sets this property.
    - `Text As String` (required): The text that appears with the Undo command.
    - `Procedure As String` (required): The name of the procedure that's run when you choose the Undo command.
- `Quit()`  
  Quits Microsoft Excel.
- `RecordMacro([BasicCode As Variant], [XlmCode As Variant])`  
  Records code if the macro recorder is on.
    - `BasicCode As Variant` (optional): A string that specifies the Visual Basic code that will be recorded if the macro recorder is recording into a Visual Basic module. The string will be recorded on one line. If the string contains a carriage return (ASCII character 10, or Chr$(10) in code), it will be recorded on more than one line.
    - `XlmCode As Variant` (optional): This argument is ignored.
- `RegisterXLL(Filename As String) As Boolean`  
  Loads an XLL code resource and automatically registers the functions and commands contained in the resource.
    - `Filename As String` (required): Specifies the name of the XLL to be loaded.
- `Repeat()`  
  Repeats the last user-interface action.
- `Undo()`  
  Cancels the last user-interface action.
- `Volatile([Volatile As Variant])`  
  Marks a user-defined function as volatile. A volatile function must be recalculated whenever calculation occurs in any cells on the worksheet. A nonvolatile function is recalculated only when the input variables change. This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell.
    - `Volatile As Variant` (optional): True to mark the function as volatile. False to mark the function as nonvolatile. The default value is True.
- `Wait(Time As Variant) As Boolean`  
  Pauses a running macro until a specified time. Returns True if the specified time has arrived.
    - `Time As Variant` (required): The time at which you want the macro to resume, in Microsoft Excel date format.
- `GetPhonetic([Text As Variant]) As String`  
  Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.
    - `Text As Variant` (optional): Specifies the text to be converted to phonetic text. If you omit this argument, the next possible phonetic text string (if any) of the previously specified _Text_ is returned. If there are no more possible phonetic text strings, an empty string is returned.
- `CalculateFull()`  
  Forces a full calculation of the data in all open workbooks.
- `FindFile() As Boolean`  
  Displays the Open dialog box.
- `CalculateFullRebuild()`  
  For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.
- `CheckAbort([KeepAbort As Variant])`  
  Stops recalculation in a Microsoft Excel application.
    - `KeepAbort As Variant` (optional): Allows recalculation to be performed for a range.
- `DisplayXMLSourcePane([XmlMap As Variant])`  
  Opens the XML Source task pane and displays the XML map specified by the _XmlMap_ argument.
    - `XmlMap As Variant` (optional): The XML map to display in the task pane.
- `CalculateUntilAsyncQueriesDone()`  
  Runs all pending queries to OLEDB and OLAP data sources.
- `SharePointVersion(bstrUrl As String) As Long`  
  Returns the version number of SharePoint Foundation instances running at the site for the specified URL.
    - `bstrUrl As String` (required): The URL of the site to check.
- `MacroOptions([Macro As Variant], [Description As Variant], [HasMenu As Variant], [MenuText As Variant], [HasShortcutKey As Variant], [ShortcutKey As Variant], [Category As Variant], [StatusBar As Variant], [HelpContextID As Variant], [HelpFile As Variant], [ArgumentDescriptions As Variant])`  
  Corresponds to options in the Macro Options dialog box. You can also use this method to display a user-defined function (UDF) in a built-in or new category within the Insert Function dialog box.
    - `Macro As Variant` (optional): The macro name or the name of a user-defined function (UDF).
    - `Description As Variant` (optional): The macro description.
    - `HasMenu As Variant` (optional): This argument is ignored.
    - `MenuText As Variant` (optional): This argument is ignored.
    - `HasShortcutKey As Variant` (optional): True to assign a shortcut key to the macro (_ShortcutKey_ must also be specified). If this argument is False, no shortcut key is assigned to the macro. If the macro already has a shortcut key, setting this argument to False removes the shortcut key. The default value is False.
    - `ShortcutKey As Variant` (optional): Required if _HasShortcutKey_ is True; ignored otherwise. The shortcut key.
    - `Category As Variant` (optional): An integer that specifies an existing macro function category (Financial, Date & Time, or User Defined, for example). See the Remarks section to determine the integers that are mapped to the built-in categories. You can also specify a string for a custom category. If you provide a string, it is treated as the category name that is displayed in the Insert Function dialog box. If the category name has never been used, a new category is defined with that name. If you use a category name that is the same as a built-in name (see list in Remarks section), Excel maps the user-defined function to that built-in category.
    - `StatusBar As Variant` (optional): The status bar text for the macro.
    - `HelpContextID As Variant` (optional): An integer that specifies the context ID for the Help topic assigned to the macro.
    - `HelpFile As Variant` (optional): The name of the Help file that contains the Help topic defined by _HelpContextId_.
    - `ArgumentDescriptions As Variant` (optional): A one-dimensional array that contains the descriptions for the arguments to a UDF that are displayed in the Function Arguments dialog box.

## Events (49)

- `NewWorkbook(Wb As Workbook)`  
  Occurs when a new workbook is created.
    - `Wb As Workbook` (required): The new workbook.
- `SheetSelectionChange(Sh As Object, Target As Range)`  
  Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).
    - `Sh As Object` (required): The worksheet that contains the new selection.
    - `Target As Range` (required): The new selected range.
- `SheetBeforeDoubleClick(Sh As Object, Target As Range, Cancel As Boolean)`  
  Occurs when any worksheet is double-clicked, before the default double-click action.
    - `Sh As Object` (required): A Worksheet object that represents the sheet.
    - `Target As Range` (required): The cell nearest to the mouse pointer when the double-click occurred.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default double-click action isn't performed when the procedure is finished.
- `SheetBeforeRightClick(Sh As Object, Target As Range, Cancel As Boolean)`  
  Occurs when any worksheet is right-clicked, before the default right-click action.
    - `Sh As Object` (required): A Worksheet object that represents the sheet.
    - `Target As Range` (required): The cell nearest to the mouse pointer when the right-click occurred.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default right-click action isn't performed when the procedure is finished.
- `SheetActivate(Sh As Object)`  
  Occurs when any sheet is activated.
    - `Sh As Object` (required): The activated sheet. Can be a Chart or Worksheet object.
- `SheetDeactivate(Sh As Object)`  
  Occurs when any sheet is deactivated.
    - `Sh As Object` (required): The sheet. Can be a Chart or Worksheet object.
- `SheetCalculate(Sh As Object)`  
  Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.
    - `Sh As Object` (required): Can be a Chart or Worksheet object.
- `SheetChange(Sh As Object, Target As Range)`  
  Occurs when cells in any worksheet are changed by the user or by an external link.
    - `Sh As Object` (required): A Worksheet object that represents the sheet.
    - `Target As Range` (required): The changed range.
- `WorkbookOpen(Wb As Workbook)`  
  Occurs when a workbook is opened.
    - `Wb As Workbook` (required): The workbook.
- `WorkbookActivate(Wb As Workbook)`  
  Occurs when any workbook is activated.
    - `Wb As Workbook` (required): The activated workbook.
- `WorkbookDeactivate(Wb As Workbook)`  
  Occurs when any open workbook is deactivated.
    - `Wb As Workbook` (required): The workbook.
- `WorkbookBeforeClose(Wb As Workbook, Cancel As Boolean)`  
  Occurs immediately before any open workbook closes.
    - `Wb As Workbook` (required): The workbook that's being closed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the workbook doesn't close when the procedure is finished.
- `WorkbookBeforeSave(Wb As Workbook, SaveAsUI As Boolean, Cancel As Boolean)`  
  Occurs before any open workbook is saved.
    - `Wb As Workbook` (required): The workbook.
    - `SaveAsUI As Boolean` (required): True if the Save As dialog box will be displayed due to changes made that need to be saved in the workbook.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the workbook isn't saved when the procedure is finished.
- `WorkbookBeforePrint(Wb As Workbook, Cancel As Boolean)`  
  Occurs before any open workbook is printed.
    - `Wb As Workbook` (required): The workbook.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the workbook isn't printed when the procedure is finished.
- `WorkbookNewSheet(Wb As Workbook, Sh As Object)`  
  Occurs when a new sheet is created in any open workbook.
    - `Wb As Workbook` (required): The workbook.
    - `Sh As Object` (required): The new sheet.
- `WorkbookAddinInstall(Wb As Workbook)`  
  Occurs when a workbook is installed as an add-in.
    - `Wb As Workbook` (required): The installed workbook.
- `WorkbookAddinUninstall(Wb As Workbook)`  
  Occurs when any add-in workbook is uninstalled.
    - `Wb As Workbook` (required): The uninstalled workbook.
- `WindowResize(Wb As Workbook, Wn As Window)`  
  Occurs when any workbook window is resized.
    - `Wb As Workbook` (required): The workbook displayed in the resized window.
    - `Wn As Window` (required): The resized window.
- `WindowActivate(Wb As Workbook, Wn As Window)`  
  Occurs when any workbook window is activated.
    - `Wb As Workbook` (required): The workbook displayed in the activated window.
    - `Wn As Window` (required): The activated window.
- `WindowDeactivate(Wb As Workbook, Wn As Window)`  
  Occurs when any workbook window is deactivated.
    - `Wb As Workbook` (required): The workbook displayed in the deactivated window.
    - `Wn As Window` (required): The deactivated window.
- `SheetFollowHyperlink(Sh As Object, Target As Hyperlink)`  
  Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the FollowHyperlink event.
    - `Sh As Object` (required): The Worksheet object that contains the hyperlink.
    - `Target As Hyperlink` (required): The Hyperlink object that represents the destination of the hyperlink.
- `SheetPivotTableUpdate(Sh As Object, Target As PivotTable)`  
  Occurs after the sheet of the PivotTable report has been updated.
    - `Sh As Object` (required): The selected sheet.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `WorkbookPivotTableCloseConnection(Wb As Workbook, Target As PivotTable)`  
  Occurs after a PivotTable report connection has been closed.
    - `Wb As Workbook` (required): The selected workbook.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `WorkbookPivotTableOpenConnection(Wb As Workbook, Target As PivotTable)`  
  Occurs after a PivotTable report connection has been opened.
    - `Wb As Workbook` (required): The selected workbook.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `WorkbookSync(Wb As Workbook, SyncEventType As MsoSyncEventType)`
- `WorkbookBeforeXmlImport(Wb As Workbook, Map As XmlMap, Url As String, IsRefresh As Boolean, Cancel As Boolean)`  
  Occurs before an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.
    - `Wb As Workbook` (required): The target workbook.
    - `Map As XmlMap` (required): The XML map that will be used to import data.
    - `Url As String` (required): The location of the XML file to be imported.
    - `IsRefresh As Boolean` (required): True if the event was triggered by refreshing an existing connection to XML data; False if a new mapping will be created.
    - `Cancel As Boolean` (required): Set to True to cancel the import or refresh operation.
- `WorkbookAfterXmlImport(Wb As Workbook, Map As XmlMap, IsRefresh As Boolean, Result As XlXmlImportResult)`  
  Occurs after an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.
    - `Wb As Workbook` (required): The target workbook.
    - `Map As XmlMap` (required): The XML map that was used to import data.
    - `IsRefresh As Boolean` (required): True if the event was triggered by refreshing an existing connection to XML data; False if a new mapping was created.
    - `Result As XlXmlImportResult` (required): Indicates the results of the refresh or import operation.
- `WorkbookBeforeXmlExport(Wb As Workbook, Map As XmlMap, Url As String, Cancel As Boolean)`  
  Occurs before Microsoft Excel saves or exports XML data from the specified workbook.
    - `Wb As Workbook` (required): The target workbook.
    - `Map As XmlMap` (required): The XML map that will be used to save or export data.
    - `Url As String` (required): The location of the XML file to be exported.
    - `Cancel As Boolean` (required): Set to True to cancel the save or export operation.
- `WorkbookAfterXmlExport(Wb As Workbook, Map As XmlMap, Url As String, Result As XlXmlExportResult)`  
  Occurs after Microsoft Excel saves or exports XML data from the specified workbook.
    - `Wb As Workbook` (required): The target workbook.
    - `Map As XmlMap` (required): The XML map that was used to save or export data.
    - `Url As String` (required): The location of the XML file that was exported.
    - `Result As XlXmlExportResult` (required): Indicates the results of the save or export operation.
- `WorkbookRowsetComplete(Wb As Workbook, Description As String, Sheet As String, Success As Boolean)`  
  The WorkbookRowsetComplete event occurs when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.
    - `Wb As Workbook` (required): The workbook for which the event occurs.
    - `Description As String` (required): A brief description of the event.
    - `Sheet As String` (required): The worksheet on which the recordset is created.
    - `Success As Boolean` (required): Contains a Boolean value to indicate success or failure.
- `AfterCalculate()`  
  The AfterCalculate event occurs when all pending refresh activity (both synchronous and asynchronous) and all of the resultant calculation activities have been completed.
- `SheetPivotTableAfterValueChange(Sh As Object, TargetPivotTable As PivotTable, TargetRange As Range)`  
  Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).
    - `Sh As Object` (required): The worksheet that contains the PivotTable.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the edited or recalculated cells.
    - `TargetRange As Range` (required): The range that contains all the edited or recalculated cells.
- `SheetPivotTableBeforeAllocateChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`  
  Occurs before changes are applied to a PivotTable.
    - `Sh As Object` (required): The worksheet that contains the PivotTable.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to apply.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList collection. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList collection. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the changes are not applied to the PivotTable, and all edits are lost.
- `SheetPivotTableBeforeCommitChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`  
  Occurs before changes are committed against the OLAP data source for a PivotTable.
    - `Sh As Object` (required): The worksheet that contains the PivotTable.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to commit.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the changes are not committed against the OLAP data source of the PivotTable.
- `SheetPivotTableBeforeDiscardChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long)`  
  Occurs before changes to a PivotTable are discarded.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to discard.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
- `ProtectedViewWindowOpen(Pvw As ProtectedViewWindow)`  
  Occurs when a workbook is opened in a Protected View window.
    - `Pvw As ProtectedViewWindow` (required): An object that represents the Protected View window that is opened.
- `ProtectedViewWindowBeforeEdit(Pvw As ProtectedViewWindow, Cancel As Boolean)`  
  Occurs immediately before editing is enabled on the workbook in the specified Protected View window.
    - `Pvw As ProtectedViewWindow` (required): The Protected View window that contains the workbook that is enabled for editing.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, editing is not enabled on the workbook.
- `ProtectedViewWindowBeforeClose(Pvw As ProtectedViewWindow, Reason As XlProtectedViewCloseReason, Cancel As Boolean)`  
  Occurs immediately before a Protected View window or a workbook in a Protected View window closes.
    - `Pvw As ProtectedViewWindow` (required): An object that represents the Protected View window that is closed.
    - `Reason As XlProtectedViewCloseReason` (required): A constant that specifies the reason the Protected View window is closed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the window does not close when the procedure is finished.
- `ProtectedViewWindowResize(Pvw As ProtectedViewWindow)`  
  Occurs when any Protected View window is resized.
    - `Pvw As ProtectedViewWindow` (required): An object that represents the resized Protected View window.
- `ProtectedViewWindowActivate(Pvw As ProtectedViewWindow)`  
  Occurs when a Protected View window is activated.
    - `Pvw As ProtectedViewWindow` (required): The activated Protected View window.
- `ProtectedViewWindowDeactivate(Pvw As ProtectedViewWindow)`  
  Occurs when a Protected View window is deactivated.
    - `Pvw As ProtectedViewWindow` (required): An object that represents the deactivated Protected View window.
- `WorkbookAfterSave(Wb As Workbook, Success As Boolean)`  
  Occurs after the workbook is saved.
    - `Wb As Workbook` (required): The workbook being saved.
    - `Success As Boolean` (required): Returns True if the save operation was successful; otherwise, False.
- `WorkbookNewChart(Wb As Workbook, Ch As Chart)`  
  Occurs when a new chart is created in any open workbook.
    - `Wb As Workbook` (required): The workbook.
    - `Ch As Chart` (required): The new chart.
- `SheetLensGalleryRenderComplete(Sh As Object)`  
  Occurs after a callout gallery's icons (dynamic and static) have finished rendering.
    - `Sh As Object` (required): Name of a worksheet.
- `SheetTableUpdate(Sh As Object, Target As TableObject)`  
  Occurs when a table on a worksheet is updated.
    - `Sh As Object` (required): The worksheet.
    - `Target As TableObject` (required): The table.
- `WorkbookModelChange(Wb As Workbook, Changes As ModelChanges)`  
  Occurs when the data model is updated.
    - `Wb As Workbook` (required): The workbook.
    - `Changes As ModelChanges` (required): The changes to the data model.
- `SheetBeforeDelete(Sh As Object)`  
  Occurs before any sheet is deleted.
    - `Sh As Object` (required): The sheet. Can be a Chart or Worksheet object.
- `WorkbookBeforeRemoteChange(Wb As Workbook)`  
  Occurs before a remote user's edits to the workbook are merged.
    - `Wb As Workbook` (required): The workbook that has been changed by a remote user.
- `WorkbookAfterRemoteChange(Wb As Workbook)`  
  Occurs after a remote user's edits to the workbook are merged.
    - `Wb As Workbook` (required): The workbook which has been changed by a remote user.
