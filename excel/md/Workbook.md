# Workbook

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020819-0000-0000-C000-000000000046}  

Represents a Microsoft Excel workbook.

**Remarks:** The Workbook object is a member of the Workbooks collection. The Workbooks collection contains all the Workbook objects currently open in Microsoft Excel. The ThisWorkbook property of the Application object returns the workbook where the Visual Basic code is running. In most cases, this is the same as the active workbook. However, if the Visual Basic code is part of an add-in, the ThisWorkbook property won't return the active workbook. In this case, the active workbook is the workbook calling the add-in, whereas the ThisWorkbook property returns the add-in workbook. If you are creating an add-in from your Visual Basic code, you should use the ThisWorkbook property to qualify any statement that must be run on the workbook that you compile into the add-in.

**Example:**

```vba
Workbooks(1).Activate
```

## Properties (121)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ActiveChart As Chart  (read-only)`  
  Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns Nothing.
- `ActiveSheet As Object  (read-only)`  
  Returns a Worksheet object that represents the active sheet (the sheet on top) in the active workbook or specified workbook. Returns Nothing if no sheet is active.
- `AutoUpdateFrequency As Long  (read/write)`  
  Returns or sets the number of minutes between automatic updates to the shared workbook. Read/write Long.
- `AutoUpdateSaveChanges As Boolean  (read/write)`  
  True if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. False if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is True. Read/write Boolean.
- `ChangeHistoryDuration As Long  (read/write)`  
  Returns or sets the number of days shown in the shared workbook's change history. Read/write Long.
- `BuiltinDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the built-in document properties for the specified workbook. Read-only.
- `Charts As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the chart sheets in the specified workbook.
- `CodeName As String  (read-only)`  
  Returns the code name for the object. Read-only String.
- `_CodeName As String  (read/write)`
- `Colors As Variant  (read/write)`  
  Returns or sets colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value. Read/write Variant.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars object that represents the Microsoft Excel command bars. Read-only.
- `ConflictResolution As XlSaveConflictResolution  (read/write)`  
  Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write XlSaveConflictResolution.
- `Container As Object  (read-only)`  
  Returns the object that represents the container application for the specified OLE object. Read-only Object.
- `CreateBackup As Boolean  (read-only)`  
  True if a backup file is created when this file is saved. Read-only Boolean.
- `CustomDocumentProperties As Object  (read-only)`  
  Returns or sets a DocumentProperties collection that represents all the custom document properties for the specified workbook.
- `Date1904 As Boolean  (read/write)`  
  True if the workbook uses the 1904 date system. Read/write Boolean.
- `DisplayDrawingObjects As XlDisplayDrawingObjects  (read/write)`  
  Returns or sets how shapes are displayed. Read/write Long.
- `FileFormat As XlFileFormat  (read-only)`  
  Returns the file format and/or type of the workbook. Read-only XlFileFormat.
- `FullName As String  (read-only)`  
  Returns the name of the object, including its path on disk, as a string. Read-only String.
- `HasPassword As Boolean  (read-only)`  
  True if the workbook has a protection password. Read-only Boolean.
- `IsAddin As Boolean  (read/write)`  
  True if the workbook is running as an add-in. Read/write Boolean.
- `Mailer As Mailer  (read-only)`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `MultiUserEditing As Boolean  (read-only)`  
  True if the workbook is open as a shared list. Read-only Boolean.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Names As Names  (read-only)`  
  Returns a Names collection that represents all the names in the specified workbook (including all worksheet-specific names). Read-only Names object.
- `Path As String  (read-only)`  
  Returns a String that represents the complete path to the workbook/file that this workbook object represents.
- `PersonalViewListSettings As Boolean  (read/write)`  
  True if filter and sort settings for lists are included in the user's personal view of the shared workbook. Read/write Boolean.
- `PersonalViewPrintSettings As Boolean  (read/write)`  
  True if print settings are included in the user's personal view of the shared workbook. Read/write Boolean.
- `PrecisionAsDisplayed As Boolean  (read/write)`  
  True if calculations in this workbook are done by using only the precision of the numbers as they're displayed. Read/write Boolean.
- `ProtectStructure As Boolean  (read-only)`  
  True if the order of the sheets in the workbook is protected. Read-only Boolean.
- `ProtectWindows As Boolean  (read-only)`  
  True if the windows of the workbook are protected. Read-only Boolean.
- `ReadOnly As Boolean  (read-only)`  
  Returns True if the object has been opened as read-only. Read-only Boolean.
- `RevisionNumber As Long  (read-only)`  
  Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero). Read-only Long.
- `Saved As Boolean  (read/write)`  
  True if no changes have been made to the specified workbook since it was last saved. Read/write Boolean.
- `SaveLinkValues As Boolean  (read/write)`  
  True if Microsoft Excel saves external link values with the workbook. Read/write Boolean.
- `Sheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the sheets in the specified workbook. Read-only Sheets object.
- `ShowConflictHistory As Boolean  (read/write)`  
  True if the Conflict History worksheet is visible in the workbook that's open as a shared list. Read/write Boolean.
- `Styles As Styles  (read-only)`  
  Returns a Styles collection that represents all the styles in the specified workbook. Read-only.
- `UpdateRemoteReferences As Boolean  (read/write)`  
  True if Microsoft Excel updates remote references in the workbook. Read/write Boolean.
- `UserStatus As Variant  (read-only)`  
  Returns a 1-based, two-dimensional array that provides information about each user who has the workbook open as a shared list. Read-only Variant.
- `CustomViews As CustomViews  (read-only)`  
  Returns a CustomViews collection that represents all the custom views for the workbook.
- `Windows As Windows  (read-only)`  
  Returns a Windows collection that represents all the windows in the specified workbook. Read-only Windows object.
- `Worksheets As Sheets  (read-only)`  
  Returns a Worksheets collection that represents all the worksheets in the specified workbook. Read-only Worksheets object.
- `WriteReserved As Boolean  (read-only)`  
  True if the workbook is write-reserved. Read-only Boolean.
- `WriteReservedBy As String  (read-only)`  
  Returns the name of the user who currently has write permission for the workbook. Read-only String.
- `Excel4IntlMacroSheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.
- `Excel4MacroSheets As Sheets  (read-only)`  
  Returns a Sheets collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.
- `TemplateRemoveExtData As Boolean  (read/write)`  
  True if external data references are removed when the workbook is saved as a template. Read/write Boolean.
- `HighlightChangesOnScreen As Boolean  (read/write)`  
  True if changes to the shared workbook are highlighted on-screen. Read/write Boolean.
- `KeepChangeHistory As Boolean  (read/write)`  
  True if change tracking is enabled for the shared workbook. Read/write Boolean.
- `ListChangesOnNewSheet As Boolean  (read/write)`  
  True if changes to the shared workbook are shown on a separate worksheet. Read/write Boolean.
- `VBProject As VBProject  (read-only)`  
  Returns a VBProject object that represents the Visual Basic project in the specified workbook. Read-only.
- `IsInplace As Boolean  (read-only)`  
  True if the specified workbook is being edited in place. False if the workbook has been opened in Microsoft Excel for editing. Read-only Boolean.
- `PublishObjects As PublishObjects  (read-only)`  
  Returns the PublishObjects collection. Read-only.
- `WebOptions As WebOptions  (read-only)`  
  Returns the WebOptions collection, which contains workbook-level attributes used by Microsoft Excel when you save a document as a webpage or open a webpage. Read-only.
- `EnvelopeVisible As Boolean  (read/write)`  
  True if the email composition header and the envelope toolbar are both visible. Read/write Boolean.
- `CalculationVersion As Long  (read-only)`  
  Returns the information about the version of Excel that the workbook was last fully recalculated by. Read-only Long.
- `VBASigned As Boolean  (read-only)`  
  True if the Visual Basic for Applications project for the specified workbook has been digitally signed. Read-only Boolean.
- `ShowPivotTableFieldList As Boolean  (read/write)`  
  True (default) if the PivotTable field list can be shown. Read/write Boolean.
- `UpdateLinks As XlUpdateLinks  (read/write)`  
  Returns or sets an XlUpdateLink constant indicating a workbook's setting for updating embedded OLE links. Read/write.
- `EnableAutoRecover As Boolean  (read/write)`  
  Saves changed files of all formats on a timed interval. Read/write Boolean.
- `RemovePersonalInformation As Boolean  (read/write)`  
  True if personal information can be removed from the specified workbook. The default value is False. Read/write Boolean.
- `FullNameURLEncoded As String  (read-only)`  
  Returns a String indicating the name of the object, including its path on disk, as a string. Read-only.
- `Password As String  (read/write)`  
  Returns or sets the password that must be supplied to open the specified workbook. Read/write String.
- `WritePassword As String  (read/write)`  
  Returns or sets a String for the write password of a workbook. Read/write.
- `PasswordEncryptionProvider As String  (read-only)`  
  Returns a String specifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.
- `PasswordEncryptionAlgorithm As String  (read-only)`  
  Returns a String indicating the algorithm that Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.
- `PasswordEncryptionKeyLength As Long  (read-only)`  
  Returns a Long indicating the key length of the algorithm that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.
- `PasswordEncryptionFileProperties As Boolean  (read-only)`  
  True if Microsoft Excel encrypts file properties for the specified password-protected workbook. Read-only Boolean.
- `ReadOnlyRecommended As Boolean  (read/write)`  
  True if the workbook was saved as read-only recommended. Read-only Boolean.
- `Permission As Permission  (read-only)`  
  Returns a Permission object that represents the permission settings in the specified workbook.
- `SharedWorkspace As SharedWorkspace  (read-only)`
- `Sync As Sync  (read-only)`
- `XmlNamespaces As XmlNamespaces  (read-only)`  
  Returns an XmlNamespaces collection that represents the XML namespaces contained in the specified workbook. Read-only.
- `XmlMaps As XmlMaps  (read-only)`  
  Returns an XmlMaps collection that represents the schema maps that have been added to the specified workbook. Read-only.
- `SmartDocument As SmartDocument  (read-only)`  
  Returns a SmartDocument object that represents the settings for a smart document solution. Read-only.
- `DocumentLibraryVersions As DocumentLibraryVersions  (read-only)`  
  Returns a DocumentLibraryVersions collection that represents the collection of versions of a shared workbook that has versioning enabled and that is stored in a document library on a server.
- `InactiveListBorderVisible As Boolean  (read/write)`  
  A Boolean value that specifies whether list borders are visible when a list is not active. Returns True if the border is visible. Read/write Boolean.
- `DisplayInkComments As Boolean  (read/write)`  
  A Boolean value that determines whether ink comments are displayed in the workbook. Read/write Boolean.
- `ContentTypeProperties As MetaProperties  (read-only)`  
  Returns a MetaProperties collection that describes the metadata stored in the workbook. Read-only.
- `Connections As Connections  (read-only)`  
  Returns a Connections object that is a container for connections between the workbook and data sources such as ODBC, OLEDB, etc., that can refresh the data without prompting the user. Read-only.
- `Signatures As SignatureSet  (read-only)`  
  Returns the digital signatures for a workbook. Read-only.
- `ServerPolicy As ServerPolicy  (read-only)`  
  Returns a ServerPolicy object that represents a policy specified for a workbook stored on a server running SharePoint Server 2007 or later. Read-only.
- `DocumentInspectors As DocumentInspectors  (read-only)`  
  Returns a DocumentInspectors collection that represents the Document Inspector modules for the specified workbook. Read-only.
- `ServerViewableItems As ServerViewableItems  (read-only)`  
  Allows a developer to interact with the list of published objects in the workbook that are shown on the server. Read-only.
- `TableStyles As TableStyles  (read-only)`  
  Returns a TableStyles collection object for the current workbook that refers to the styles used in the current workbook. Read-only.
- `DefaultTableStyle As Variant  (read/write)`  
  Specifies the table style from the TableStyles collection that is used as the default table style. Read/write Variant.
- `DefaultPivotTableStyle As Variant  (read/write)`  
  Specifies the table style from the TableStyles collection that is used as the default style for PivotTables. Read/write.
- `CheckCompatibility As Boolean  (read/write)`  
  Controls whether or not the compatibility checker is run automatically when the workbook is saved. Read/write Boolean.
- `HasVBProject As Boolean  (read-only)`  
  Returns a Boolean that represents whether a workbook has an attached Microsoft Visual Basic for Applications project. Read-only Boolean.
- `CustomXMLParts As CustomXMLParts  (read-only)`  
  Returns a CustomXMLParts collection that represents the custom XML in the XML data store. Read-only.
- `Final As Boolean  (read/write)`  
  Returns or sets a Boolean that indicates whether a workbook is final. Read/write Boolean.
- `Research As Research  (read-only)`  
  Returns a Research object that represents the research service for a workbook. Read-only.
- `Theme As OfficeTheme  (read-only)`  
  Returns the theme applied to the current workbook. Read-only.
- `Excel8CompatibilityMode As Boolean  (read-only)`  
  Provides developers with a way to check if the workbook is in compatibility mode. Read-only Boolean.
- `ConnectionsDisabled As Boolean  (read-only)`  
  Disables the external connections or links in the workbook. Read-only.
- `ShowPivotChartActiveFields As Boolean  (read/write)`  
  This property controls the visibility of the PivotChart Filter pane. Read/write Boolean.
- `IconSets As IconSets  (read-only)`  
  This property is used to filter data in a workbook based on a cell icon from the IconSets collection. Read-only.
- `EncryptionProvider As String  (read/write)`  
  Returns a String specifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting documents. Read/write.
- `DoNotPromptForConvert As Boolean  (read/write)`  
  Returns or sets if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Read/write Boolean.
- `ForceFullCalculation As Boolean  (read/write)`  
  Returns or sets the specified workbook to forced calculation mode. Read/write.
- `SlicerCaches As SlicerCaches  (read-only)`  
  Returns the SlicerCaches object associated with the workbook. Read-only.
- `ActiveSlicer As Slicer  (read-only)`  
  Returns an object that represents the active slicer in the active workbook or in the specified workbook. Returns Nothing if no slicer is active. Read-only.
- `DefaultSlicerStyle As Variant  (read/write)`  
  Specifies the style from the TableStyles object that is used as the default style for slicers. Read/write.
- `AccuracyVersion As Long  (read/write)`  
  Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Read/write.
- `CaseSensitive As Boolean  (read-only)`  
  True if the workbook distinguishes between uppercase and lowercase when comparing content. Read-only Boolean.
- `UseWholeCellCriteria As Boolean  (read-only)`  
  True if the workbook uses search patterns that match the entire content of a cell. Read-only Boolean.
- `UseWildcards As Boolean  (read-only)`  
  True if the workbook enables wildcards for character string comparisons and searching. Read-only Boolean.
- `PivotTables As Object  (read-only)`  
  Returns an object that represents a collection of all the PivotTable reports on a worksheet. Read-only.
- `Model As Model  (read-only)`  
  Returns the top-level Model object that is the one data model for the workbook. Read-only.
- `ChartDataPointTrack As Boolean  (read/write)`  
  True causes all charts in the current document to track the actual data point to which it's attached. False reverts back to tracking the index of the data point. Read/write Boolean.
- `DefaultTimelineStyle As Variant  (read/write)`  
  The name of the default slicer style of the workbook. Read/write Variant.
- `Queries As Queries  (read-only)`  
  Returns a Queries collection that represents all the Get & Transform queries in the specified Workbook. Starting from Excel 2016, Get & Transform features enable you to connect, combine, and shape data from a variety of sources to meet your analysis needs. Read-only.
- `WorkIdentity As String  (read/write)`
- `AutoSaveOn As Boolean  (read/write)`  
  True if the edits in the workbook are automatically saved. Read/write Boolean.
- `SensitivityLabel As ISensitivityLabel  (read-only)`  
  Returns the Microsoft Office SensitivityLabel object from the Workbook.
- `ExternalCodeServiceTimeout As Long  (read/write)`
- `CompatibilityVersion As Long  (read/write)`

## Methods (65)

- `Activate()`  
  Activates the first window associated with the workbook.
- `ChangeFileAccess(Mode As XlFileAccess, [WritePassword As Variant], [Notify As Variant])`  
  Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.
    - `Mode As XlFileAccess` (required): Specifies the new access mode.
    - `WritePassword As Variant` (optional): Specifies the write-reserved password if the file is write reserved and _Mode_ is xlReadWrite. Ignored if there's no password for the file or if _Mode_ is xlReadOnly.
    - `Notify As Variant` (optional): True (or omitted) to notify the user if the file cannot be immediately accessed.
- `ChangeLink(Name As String, NewName As String, [Type As XlLinkType])`  
  Changes a link from one document to another.
    - `Name As String` (required): The name of the Microsoft Excel or DDE/OLE link to be changed, as it was returned from the LinkSources method.
    - `NewName As String` (required): The new name of the link.
    - `Type As XlLinkType` (optional): The link type.
- `Close([SaveChanges As Variant], [Filename As Variant], [RouteWorkbook As Variant])`  
  Closes the object.
    - `SaveChanges As Variant` (optional): If there are no changes to the workbook, this argument is ignored. If there are changes to the workbook and the workbook appears in other open windows, this argument is ignored. If there are changes to the workbook but the workbook doesn't appear in any other open windows, this argument specifies whether changes should be saved. If set to True, changes are saved to the workbook. If there is not yet a file name associated with the workbook, _FileName_ is used. If _FileName_ is omitted, the user is asked to supply a file name.
    - `Filename As Variant` (optional): Saves changes under this file name.
    - `RouteWorkbook As Variant` (optional): If the workbook doesn't need to be routed to the next recipient (if it has no routing slip or has already been routed), this argument is ignored. Otherwise, Microsoft Excel routes the workbook according to the value of this parameter. If set to True, the workbook is sent to the next recipient. If set to False, the workbook is not sent. If omitted, the user is asked whether the workbook should be sent.
- `DeleteNumberFormat(NumberFormat As String)`  
  Deletes a custom number format from the workbook.
    - `NumberFormat As String` (required): Names the number format to be deleted.
- `ExclusiveAccess() As Boolean`  
  Assigns the current user exclusive access to the workbook that's open as a shared list.
- `ForwardMailer()`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `LinkInfo(Name As String, LinkInfo As XlLinkInfo, [Type As Variant], [EditionRef As Variant]) As Variant`  
  Returns the link date and update status.
    - `Name As String` (required): The name of the link.
    - `LinkInfo As XlLinkInfo` (required): The type of information to be returned.
    - `Type As Variant` (optional): One of the constants of XlLinkInfoType specifying the type of link to return.
    - `EditionRef As Variant` (optional): If the link is an edition, this argument specifies the edition reference as a string in R1C1 style. This argument is required if there's more than one publisher or subscriber with the same name in the workbook.
- `LinkSources([Type As Variant]) As Variant`  
  Returns an array of links in the workbook. The names in the array are the names of the linked documents, editions, or DDE or OLE servers. Returns Empty if there are no links.
    - `Type As Variant` (optional): One of the constants of XlLink, which specifies the type of link to return.
- `MergeWorkbook(Filename As Variant)`  
  Merges changes from one workbook into an open workbook.
    - `Filename As Variant` (required): The file name of the workbook that contains the changes to be merged into the open workbook.
- `NewWindow() As Window`  
  Creates a new window or a copy of the specified window.
- `OpenLinks(Name As String, [ReadOnly As Variant], [Type As Variant])`  
  Opens the supporting documents for a link or links.
    - `Name As String` (required): The name of the Microsoft Excel or DDE/OLE link, as returned from the LinkSources method.
    - `ReadOnly As Variant` (optional): True to open documents as read-only. The default value is False.
    - `Type As Variant` (optional): One of the constants of XlLink that specifies the link type.
- `PivotCaches() As PivotCaches`  
  Returns a PivotCaches collection that represents all the PivotTable caches in the specified workbook. Read-only.
- `Post([DestName As Variant])`  
  Posts the specified workbook to a public folder. This method works only with a Microsoft Exchange client connected to a Microsoft Exchange server.
    - `DestName As Variant` (optional): This argument is ignored. The Post method prompts the user to specify the destination for the workbook.
- `PrintPreview([EnableChanges As Variant])`  
  Shows a preview of the object as it would look when printed.
    - `EnableChanges As Variant` (optional): Pass a Boolean value to specify if the user can change the margins and other page setup options available in print preview.
- `RefreshAll()`  
  Refreshes all external data ranges and PivotTable reports in the specified workbook.
- `Reply()`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `ReplyAll()`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.
- `RemoveUser(Index As Long)`  
  Disconnects the specified user from the shared workbook.
    - `Index As Long` (required): The user index.
- `RunAutoMacros(Which As XlRunAutoMacro)`  
  Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook. This method is included for backward compatibility. For new Visual Basic code, you should use the Open, Activate, and Deactivate events and the Close method instead of these macros.
    - `Which As XlRunAutoMacro` (required): Specifies the automatic macro to run.
- `Save()`  
  Saves changes to the specified workbook.
- `SaveCopyAs([Filename As Variant])`  
  Saves a copy of the workbook to a file but doesn't modify the open workbook in memory.
    - `Filename As Variant` (optional): Specifies the file name for the copy.
- `SendMail(Recipients As Variant, [Subject As Variant], [ReturnReceipt As Variant])`  
  Sends the workbook by using the installed mail system.
    - `Recipients As Variant` (required): Specifies the name of the recipient as text, or as an array of text strings if there are multiple recipients. At least one recipient must be specified, and all recipients are added as To recipients.
    - `Subject As Variant` (optional): Specifies the subject of the message. If this argument is omitted, the document name is used.
    - `ReturnReceipt As Variant` (optional): True to request a return receipt. False to not request a return receipt. The default value is False.
- `SendMailer([FileFormat As Variant], [Priority As XlPriority])`  
  You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.
    - `FileFormat As Variant` (optional): See help included with Microsoft Office Macintosh Edition.
    - `Priority As XlPriority` (optional): See help included with Microsoft Office Macintosh Edition.
- `SetLinkOnData(Name As String, [Procedure As Variant])`  
  Sets the name of a procedure that runs whenever a DDE link is updated.
    - `Name As String` (required): The name of the DDE/OLE link, as returned from the LinkSources method.
    - `Procedure As Variant` (optional): The name of the procedure to be run when the link is updated. This can be either a Microsoft Excel 4.0 macro or a Visual Basic procedure. Set this argument to an empty string ("") to indicate that no procedure should run when the link is updated.
- `Unprotect([Password As Variant])`  
  Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
    - `Password As Variant` (optional): A string that denotes the case-sensitive password to use to unprotect the sheet or workbook. If the sheet or workbook isn't protected with a password, this argument is ignored. If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. If you omit this argument for a workbook that's protected with a password, the method fails.
- `UnprotectSharing([SharingPassword As Variant])`  
  Turns off protection for sharing and saves the workbook.
    - `SharingPassword As Variant` (optional): The workbook password.
- `UpdateFromFile()`  
  Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory. If the disk copy hasn't changed since the workbook was loaded, the in-memory copy of the workbook isn't reloaded.
- `UpdateLink([Name As Variant], [Type As Variant])`  
  Updates a Microsoft Excel, DDE, or OLE link (or links).
    - `Name As Variant` (optional): The name of the Microsoft Excel or DDE/OLE link to be updated, as returned from the LinkSources method.
    - `Type As Variant` (optional): One of the constants of XlLinkType specifying the type of link.
- `HighlightChangesOptions([When As Variant], [Who As Variant], [Where As Variant])`  
  Controls how changes are shown in a shared workbook.
    - `When As Variant` (optional): The changes that are shown. Can be one of the following XlHighlightChangesTime constants: xlSinceMyLastSave, xlAllChanges, or xlNotYetReviewed.
    - `Who As Variant` (optional): The user or users whose changes are shown. Can be "Everyone," "Everyone but Me," or the name of one of the users of the shared workbook.
    - `Where As Variant` (optional): An A1-style range reference that specifies the area to check for changes.
- `PurgeChangeHistoryNow(Days As Long, [SharingPassword As Variant])`  
  Removes entries from the change log for the specified workbook.
    - `Days As Long` (required): The number of days that changes in the change log are to be retained.
    - `SharingPassword As Variant` (optional): The password that unprotects the workbook for sharing. If the workbook is protected for sharing with a password and this argument is omitted, the user is prompted for the password.
- `AcceptAllChanges([When As Variant], [Who As Variant], [Where As Variant])`  
  Accepts all changes in the specified shared workbook.
    - `When As Variant` (optional): Specifies when all the changes are accepted.
    - `Who As Variant` (optional): Specifies by whom all the changes are accepted.
    - `Where As Variant` (optional): Specifies where all the changes are accepted.
- `RejectAllChanges([When As Variant], [Who As Variant], [Where As Variant])`  
  Rejects all changes in the specified shared workbook.
    - `When As Variant` (optional): Specifies when all the changes are rejected.
    - `Who As Variant` (optional): Specifies by whom all the changes are rejected.
    - `Where As Variant` (optional): Specifies where all the changes are rejected.
- `ResetColors()`  
  Resets the color palette to the default colors.
- `FollowHyperlink(Address As String, [SubAddress As Variant], [NewWindow As Variant], [AddHistory As Variant], [ExtraInfo As Variant], [Method As Variant], [HeaderInfo As Variant])`  
  Displays a cached document if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
    - `Address As String` (required): The address of the target document.
    - `SubAddress As Variant` (optional): The location within the target document. The default value is the empty string.
    - `NewWindow As Variant` (optional): True to display the target application in a new window. The default value is False.
    - `AddHistory As Variant` (optional): Not used. Reserved for future use.
    - `ExtraInfo As Variant` (optional): A String or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use _ExtraInfo_ to specify the coordinates of an image map, the contents of a form, or a FAT file name.
    - `Method As Variant` (optional): Specifies the way _ExtraInfo_ is attached. Can be one of the MsoExtraInfoMethod constants: msoMethodGet or msoMethodPost.
    - `HeaderInfo As Variant` (optional): A String that specifies header information for the HTTP request. The default value is an empty string.
- `AddToFavorites()`  
  Adds a shortcut to the workbook or hyperlink to the Favorites folder.
- `WebPagePreview()`  
  Displays a preview of the specified workbook as it would look if saved as a webpage.
- `ReloadAs(Encoding As MsoEncoding)`  
  Reloads a workbook based on an HTML document, using the specified document encoding.
    - `Encoding As MsoEncoding` (required): The encoding that is to be applied to the workbook.
- `BreakLink(Name As String, Type As XlLinkType)`  
  Converts formulas linked to other Microsoft Excel sources or OLE sources to values.
    - `Name As String` (required): The name of the link.
    - `Type As XlLinkType` (required): The type of link.
- `CheckIn([SaveChanges As Variant], [Comments As Variant], [MakePublic As Variant])`  
  Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.
    - `SaveChanges As Variant` (optional): True saves changes and checks in the document. False returns the document to a checked-in status without saving revisions.
    - `Comments As Variant` (optional): Allows the user to enter check-in comments for the revision of the workbook being checked in (applies only if _SaveChanges_ equals True).
    - `MakePublic As Variant` (optional): True allows the user to publish the workbook after it has been checked in. This submits the workbook for the approval process, which can eventually result in a version of the workbook being published to users with read-only rights to the workbook (applies only if _SaveChanges_ equals True).
- `CanCheckIn() As Boolean`  
  True if Microsoft Excel can check in a specified workbook to a server. Read/write Boolean.
- `SendForReview([Recipients As Variant], [Subject As Variant], [ShowMessage As Variant], [IncludeAttachment As Variant])`  
  Sends a workbook in an email message for review to the specified recipients.
    - `Recipients As Variant` (optional): A string that lists the people to whom to send the message. These can be unresolved names and aliases in an email phone book or full email addresses. Separate multiple recipients with a semicolon (;). If left blank and _ShowMessage_ is False, you'll receive an error message, and the message will not be sent.
    - `Subject As Variant` (optional): A string for the subject of the message. If left blank, the subject will be: Please review "filename".
    - `ShowMessage As Variant` (optional): A Boolean value that indicates whether the message should be displayed when the method is executed. The default value is True. If set to False, the message is automatically sent to the recipients without first showing the message to the sender.
    - `IncludeAttachment As Variant` (optional): A Boolean value that indicates whether the message should include an attachment or a link to a server location. The default value is True. If set to False, the document must be stored at a shared location.
- `ReplyWithChanges([ShowMessage As Variant])`  
  Sends an email message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.
    - `ShowMessage As Variant` (optional): False does not display the message. True displays the message.
- `EndReview()`  
  Terminates a review of a file that has been sent for review by using the SendForReview method.
- `SetPasswordEncryptionOptions([PasswordEncryptionProvider As Variant], [PasswordEncryptionAlgorithm As Variant], [PasswordEncryptionKeyLength As Variant], [PasswordEncryptionFileProperties As Variant])`  
  Sets the options for encrypting workbooks by using passwords.
    - `PasswordEncryptionProvider As Variant` (optional): A case-sensitive string of the encryption provider.
    - `PasswordEncryptionAlgorithm As Variant` (optional): A case-sensitive string of the algorithmic short name (that is, "RC4").
    - `PasswordEncryptionKeyLength As Variant` (optional): The encryption key length which is a multiple of 8 (40 or greater).
    - `PasswordEncryptionFileProperties As Variant` (optional): True (default) to encrypt file properties.
- `Protect([Password As Variant], [Structure As Variant], [Windows As Variant])`  
  Protects a workbook so that it cannot be modified.
    - `Password As Variant` (optional): A string that specifies a case-sensitive password for the worksheet or workbook. If this argument is omitted, you can unprotect the worksheet or workbook without using a password. Otherwise, you must specify the password to unprotect the worksheet or workbook. If you forget the password, you cannot unprotect the worksheet or workbook. Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. It's critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect.
    - `Structure As Variant` (optional): True to protect the structure of the workbook (the relative position of the sheets). The default value is False.
    - `Windows As Variant` (optional): True to protect the workbook windows. If this argument is omitted, the windows aren't protected.
- `SendFaxOverInternet([Recipients As Variant], [Subject As Variant], [ShowMessage As Variant])`  
  Sends a worksheet as a fax to the specified recipients.
    - `Recipients As Variant` (optional): A String that represents the fax numbers and email addresses of the people to whom the fax will be sent. Separate multiple recipients with a semicolon.
    - `Subject As Variant` (optional): A String that represents the subject line for the faxed document.
    - `ShowMessage As Variant` (optional): True displays the fax message before sending it. False sends the fax without displaying the fax message.
- `XmlImport(Url As String, ImportMap As XmlMap, [Overwrite As Variant], [Destination As Variant]) As XlXmlImportResult`  
  Imports an XML data file into the current workbook.
    - `Url As String` (required): A uniform resource locator (URL) or a uniform naming convention (UNC) path to an XML data file.
    - `ImportMap As XmlMap` (required): The schema map to apply when importing the file. If the data was previously imported, contains a reference to the XmlMap object containing the data.
    - `Overwrite As Variant` (optional): If a value is not specified for the _Destination_ parameter, this parameter specifies whether to overwrite data that has been mapped to the schema map specified in the _ImportMap_ parameter. Set to True to overwrite the data or False to append the new data to the existing data. The default value is True. If a value is specified for the _Destination_ parameter, this parameter specifies whether to overwrite existing data. Set to True to overwrite existing data or False to cancel the import if data would be overwritten. The default value is True.
    - `Destination As Variant` (optional): Specifies the range where the list will be created. You only use the top-left corner of the range.
- `XmlImportXml(Data As String, ImportMap As XmlMap, [Overwrite As Variant], [Destination As Variant]) As XlXmlImportResult`  
  Imports an XML data stream that has been previously loaded into memory. Excel uses the first qualifying map found, or if the destination range is specified, Excel automatically lists the data.
    - `Data As String` (required): The data to import.
    - `ImportMap As XmlMap` (required): The schema map to apply when importing the file.
    - `Overwrite As Variant` (optional): If a value is not specified for the _Destination_ parameter, this parameter specifies whether to overwrite data that has been mapped to the schema map specified in the _ImportMap_ parameter. Set to True to overwrite the data or False to append the new data to the existing data. The default value is True. If a value is specified for the _Destination_ parameter, this parameter specifies whether to overwrite existing data. Set to True to overwrite existing data or False to cancel the import if data would be overwritten. The default value is True.
    - `Destination As Variant` (optional): Specifies the range where the list will be created. Excel only uses the top-left corner of the range.
- `SaveAsXMLData(Filename As String, Map As XmlMap)`  
  Exports the data that has been mapped to the specified XML schema map to an XML data file.
    - `Filename As String` (required): A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.
    - `Map As XmlMap` (required): The schema map to apply to the data.
- `ToggleFormsDesign()`  
  Used to toggle Excel into Design mode when using forms controls.
- `RemoveDocumentInformation(RemoveDocInfoType As XlRemoveDocInfoType)`  
  Removes all information of the specified type from the workbook.
    - `RemoveDocInfoType As XlRemoveDocInfoType` (required): Type of information to be removed.
- `CheckInWithVersion([SaveChanges As Variant], [Comments As Variant], [MakePublic As Variant], [VersionType As Variant])`  
  Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally.
    - `SaveChanges As Variant` (optional): True to save the workbook to the server location. The default is True.
    - `Comments As Variant` (optional): Comments for the revision of the workbook being checked in (applies only if _SaveChanges_ is set to True).
    - `MakePublic As Variant` (optional): True to allow the user to publish the workbook after it is checked in.
    - `VersionType As Variant` (optional): Specifies versioning information for the workbook.
- `LockServerFile()`  
  Locks the workbook on the server to prevent modification.
- `GetWorkflowTasks() As WorkflowTasks`  
  Returns the collection of WorkflowTask objects for the specified workbook.
- `GetWorkflowTemplates() As WorkflowTemplates`  
  Returns the collection of WorkflowTemplate objects for the specified workbook.
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant], [IgnorePrintAreas As Variant])`  
  Prints the object.
    - `From As Variant` (optional): The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to print. If this argument is omitted, printing ends with the last page.
    - `Copies As Variant` (optional): The number of copies to print. If this argument is omitted, one copy is printed.
    - `Preview As Variant` (optional): True to have Microsoft Excel invoke print preview before printing the object. False (or omitted) to print the object immediately.
    - `ActivePrinter As Variant` (optional): Sets the name of the active printer.
    - `PrintToFile As Variant` (optional): True to print to a file. If _PrToFileName_ is not specified, Microsoft Excel prompts the user to enter the name of the output file.
    - `Collate As Variant` (optional): True to collate multiple copies.
    - `PrToFileName As Variant` (optional): If _PrintToFile_ is set to True, this argument specifies the name of the file that you want to print to.
    - `IgnorePrintAreas As Variant` (optional): True to ignore print areas and print the entire object.
- `ApplyTheme(Filename As String)`  
  Applies the specified theme to the current workbook.
    - `Filename As String` (required): Full path and file name of a stored theme
- `EnableConnections()`  
  The EnableConnections method allows developers to programmatically enable data connections within the workbook for the user.
- `ProtectSharing([Filename As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [SharingPassword As Variant], [FileFormat As Variant])`  
  Saves the workbook and protects it for sharing.
    - `Filename As Variant` (optional): A string indicating the name of the saved file. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.
    - `Password As Variant` (optional): A case-sensitive string indicating the protection password to be given to the file. Should be no longer than 15 characters.
    - `WriteResPassword As Variant` (optional): A string indicating the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened read-only.
    - `ReadOnlyRecommended As Variant` (optional): True to display a message when the file is opened, recommending that the file be opened read-only.
    - `CreateBackup As Variant` (optional): True to create a backup file.
    - `SharingPassword As Variant` (optional): A string indicating the password to be used to protect the file for sharing.
    - `FileFormat As Variant` (optional): A string indicating the file format.
- `CreateForecastSheet(Timeline As Range, Values As Range, [ForecastStart As Variant], [ForecastEnd As Variant], [ConfInt As Variant], [Seasonality As Variant], [DataCompletion As Variant], [Aggregation As Variant], [ChartType As Variant], [ShowStatsTable As Variant])`  
  If you have historical time-based data, you can use CreateForecastSheet to create a forecast. When you create a forecast, a new worksheet is created that contains a table of the historical and predicted values and a chart showing this. A forecast can help you predict things like future sales, inventory requirements, or consumer trends.
- `SaveAs([Filename As Variant], [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AccessMode As XlSaveAsAccessMode], [ConflictResolution As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant], [WorkIdentity As Variant])`  
  Saves changes to the workbook in a different file.
    - `Filename As Variant` (optional): A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.
    - `FileFormat As Variant` (optional): The file format to use when you save the file. For a list of valid choices, see the XlFileFormat enumeration. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.
    - `Password As Variant` (optional): A case-sensitive string (no more than 15 characters) that indicates the protection password to be given to the file.
    - `WriteResPassword As Variant` (optional): A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only.
    - `ReadOnlyRecommended As Variant` (optional): True to display a message when the file is opened, recommending that the file be opened as read-only.
    - `CreateBackup As Variant` (optional): True to create a backup file.
    - `AccessMode As XlSaveAsAccessMode` (optional): The access mode for the workbook.
    - `ConflictResolution As Variant` (optional): An XlSaveConflictResolution value that determines how the method resolves a conflict while saving the workbook. If set to xlUserResolution, the conflict-resolution dialog box is displayed. If set to xlLocalSessionChanges, the local user's changes are automatically accepted. If set to xlOtherSessionChanges, the changes from other sessions are automatically accepted instead of the local user's changes. If this argument is omitted, the conflict-resolution dialog box is displayed.
    - `AddToMru As Variant` (optional): True to add this workbook to the list of recently used files. The default value is False.
    - `TextCodepage As Variant` (optional): Ignored for all languages in Microsoft Excel. NOTE: When Excel saves a workbook to one of the CSV or text formats, which are specified by using the _FileFormat_ parameter, it uses the code page that corresponds to the language for the system locale in use on the current computer. This system setting is available in the Control Panel > Region and Language > Location tab under Current location.
    - `TextVisualLayout As Variant` (optional): Ignored for all languages in Microsoft Excel. NOTE: When Excel saves a workbook to one of the CSV or text formats, which are specified by using the _FileFormat_ parameter, it saves these formats in logical layout. If left-to-right (LTR) text is embedded within right-to-left (RTL) text in the file, or vice versa, logical layout saves the contents of the file in the correct reading order for all languages in the file without regard to direction. When an application opens the file, each run of LTR or RTL characters are rendered in the correct direction according to the character value ranges within the code page (unless an application that is designed to display the exact memory layout of the file, such as a debugger or editor, is used to open the file).
    - `Local As Variant` (optional): True saves files against the language of Microsoft Excel (including control panel settings). False (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically US English unless the VBA project where Workbooks.Open is run from is an old internationalized XL5/95 VBA project).
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`  
  The ExportAsFixedFormat method is used to publish a workbook to either the PDF or XPS format.
    - `Type As XlFixedFormatType` (required): Can be either xlTypePDF or xlTypeXPS.
    - `Filename As Variant` (optional): A string that indicates the name of the file to be saved. You can include a full path, or Excel saves the file in the current folder.
    - `Quality As Variant` (optional): Can be set to either of the following XlFixedFormatQuality constants: xlQualityStandard or xlQualityMinimum.
    - `IncludeDocProperties As Variant` (optional): Set to True to indicate that document properties should be included, or set to False to indicate that they are omitted.
    - `IgnorePrintAreas As Variant` (optional): If set to True, ignores any print areas set when publishing. If set to False, uses the print areas set when publishing.
    - `From As Variant` (optional): The number of the page at which to start publishing. If this argument is omitted, publishing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to publish. If this argument is omitted, publishing ends with the last page.
    - `OpenAfterPublish As Variant` (optional): If set to True, displays the file in the viewer after it is published. If set to False, the file is published but not displayed.
    - `FixedFormatExtClassPtr As Variant` (optional): Pointer to the FixedFormatExt class.
- `PublishToPBI([PublishType As Variant], [nameConflict As Variant], [bstrGroupName As Variant]) As String`
- `ConvertComments()`  
  Converts all legacy comments and notes to modern comments.

## Events (42)

- `Open()`  
  Occurs when the workbook is opened.
- `Activate()`  
  Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.
- `Deactivate()`  
  Occurs when the chart, worksheet, or workbook is deactivated.
- `BeforeClose(Cancel As Boolean)`  
  Occurs before the workbook closes. If the workbook has been changed, this event occurs before the user is asked to save changes.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the close operation stops and the workbook is left open.
- `BeforeSave(SaveAsUI As Boolean, Cancel As Boolean)`  
  Occurs before the workbook is saved.
    - `SaveAsUI As Boolean` (required): True if the Save As dialog box is displayed due to changes made that need to be saved in the workbook.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the workbook isn't saved when the procedure is finished.
- `BeforePrint(Cancel As Boolean)`  
  Occurs before the workbook (or anything in it) is printed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the workbook isn't printed when the procedure is finished.
- `NewSheet(Sh As Object)`  
  Occurs when a new sheet is created in the workbook.
    - `Sh As Object` (required): The new sheet. Can be a Worksheet or Chart object.
- `AddinInstall()`  
  Occurs when the workbook is installed as an add-in.
- `AddinUninstall()`  
  Occurs when the workbook is uninstalled as an add-in.
- `WindowResize(Wn As Window)`  
  Occurs when any workbook window is resized.
    - `Wn As Window` (required): The resized window.
- `WindowActivate(Wn As Window)`  
  Occurs when any workbook window is activated.
    - `Wn As Window` (required): The activated window.
- `WindowDeactivate(Wn As Window)`  
  Occurs when any workbook window is deactivated.
    - `Wn As Window` (required): The deactivated window.
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
- `SheetFollowHyperlink(Sh As Object, Target As Hyperlink)`  
  Occurs when you choose any hyperlink in Microsoft Excel. For worksheet-level events, see the FollowHyperlink event.
    - `Sh As Object` (required): The Worksheet object that contains the hyperlink.
    - `Target As Hyperlink` (required): The Hyperlink object that represents the destination of the hyperlink.
- `SheetPivotTableUpdate(Sh As Object, Target As PivotTable)`  
  Occurs after the sheet of the PivotTable report has been updated.
    - `Sh As Object` (required): The selected sheet.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `PivotTableCloseConnection(Target As PivotTable)`  
  Occurs after a PivotTable report closes the connection to its data source.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `PivotTableOpenConnection(Target As PivotTable)`  
  Occurs after a PivotTable report opens the connection to its data source.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `Sync(SyncEventType As MsoSyncEventType)`
- `BeforeXmlImport(Map As XmlMap, Url As String, IsRefresh As Boolean, Cancel As Boolean)`  
  Occurs before an existing XML data connection is refreshed or before new XML data is imported into a Microsoft Excel workbook.
    - `Map As XmlMap` (required): The XML map that is used to import data.
    - `Url As String` (required): The location of the XML file to be imported.
    - `IsRefresh As Boolean` (required): True if the event was triggered by refreshing an existing connection to XML data; False if the event was triggered by importing from a different data source.
    - `Cancel As Boolean` (required): Set to True to cancel the import or refresh operation.
- `AfterXmlImport(Map As XmlMap, IsRefresh As Boolean, Result As XlXmlImportResult)`  
  Occurs after an existing XML data connection is refreshed or after new XML data is imported into the specified Microsoft Excel workbook.
    - `Map As XmlMap` (required): The XML map that will be used to import data.
    - `IsRefresh As Boolean` (required): True if the event was triggered by refreshing an existing connection to XML data; False if the event was triggered by importing from a different data source.
    - `Result As XlXmlImportResult` (required): Indicates the results of the refresh or import operation.
- `BeforeXmlExport(Map As XmlMap, Url As String, Cancel As Boolean)`  
  Occurs before Microsoft Excel saves or exports XML data from the specified workbook.
    - `Map As XmlMap` (required): The XML map that is used to save or export data.
    - `Url As String` (required): The location where you want to export the resulting XML file.
    - `Cancel As Boolean` (required): Set to True to cancel the save or export operation.
- `AfterXmlExport(Map As XmlMap, Url As String, Result As XlXmlExportResult)`  
  Occurs after Microsoft Excel saves or exports XML data from the specified workbook.
    - `Map As XmlMap` (required): The schema map that was used to save or export data.
    - `Url As String` (required): The location of the XML file that was exported.
    - `Result As XlXmlExportResult` (required): Indicates the results of the save or export operation.
- `RowsetComplete(Description As String, Sheet As String, Success As Boolean)`  
  The event is raised when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.
    - `Description As String` (required): A brief description of the event.
    - `Sheet As String` (required): Worksheet on which the recordset is created.
    - `Success As Boolean` (required): Contains a Boolean value to indicate success or failure.
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
    - `Sh As Object` (required): The worksheet that contains the PivotTable.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to discard.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
- `SheetPivotTableChangeSync(Sh As Object, Target As PivotTable)`  
  Occurs after changes to a PivotTable.
    - `Sh As Object` (required): The worksheet that contains the PivotTable.
    - `Target As PivotTable` (required): The PivotTable that was changed.
- `AfterSave(Success As Boolean)`  
  Occurs after the workbook is saved.
    - `Success As Boolean` (required): True if the save operation was successful; False otherwise.
- `NewChart(Ch As Chart)`  
  Occurs when a new chart is created in the workbook.
    - `Ch As Chart` (required): The new chart.
- `SheetLensGalleryRenderComplete(Sh As Object)`  
  Occurs when a callout gallery's icons (dynamic and static) have completed rendering for a worksheet.
    - `Sh As Object` (required): A worksheet object.
- `SheetTableUpdate(Sh As Object, Target As TableObject)`  
  Occurs after the sheet table has been updated.
    - `Sh As Object` (required): The selected sheet.
    - `Target As TableObject` (required): The selected table.
- `ModelChange(Changes As ModelChanges)`  
  Occurs after the Excel data model is changed.
    - `Changes As ModelChanges` (required): ModelChanges object that includes changes made to the Excel data model during the last transaction.
- `SheetBeforeDelete(Sh As Object)`  
  Occurs when any sheet is deleted.
    - `Sh As Object` (required): The sheet. Can be a Chart or Worksheet object.
- `BeforeRemoteChange()`  
  Occurs before a remote user's edits to the workbook are merged.
- `AfterRemoteChange()`  
  Occurs after a remote user's edits to the workbook are merged.
