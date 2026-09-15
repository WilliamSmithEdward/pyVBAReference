# Application

**Type:** Class  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209FF-0000-0000-C000-000000000046}  

Represents the Microsoft Word application. The Application object includes properties and methods that return top-level objects. For example, the ActiveDocument property returns a Document object.

**Remarks:** Use the Application property to return the Application object. The following example displays the user name for Word. Many of the properties and methods that return the most common user-interface objects-such as the active document (ActiveDocument property)-can be used without the Application object qualifier. For example, instead of writing Application.ActiveDocument.PrintOut, you can write ActiveDocument.PrintOut. Properties and methods that can be used without the Application object qualifier are considered "global." To view the global properties and methods in the Object Browser, click <globals> at the top of the list in the Classes box. (Also see the Global object.) Remarks To use Automation (formerly OLE Automation) to control Word from another application, use the Microsoft Visual Basic CreateObject or GetObject function to return a Word Application object. The following Microsoft Excel example starts Word (if it is not already running) and opens an existing document.

## Properties (121)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Application object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Documents As Documents  (read-only)`  
  Returns a Documents collection that represents all the open documents. Read-only.
- `Windows As Windows  (read-only)`  
  Returns a Windows collection that represents all document windows. Read-only.
- `ActiveDocument As Document  (read-only)`  
  Returns a Document object that represents the active document (the document with the focus). If there are no documents open, an error occurs. Read-only.
- `ActiveWindow As Window  (read-only)`  
  Returns a Window object that represents the active window (the window with the focus). If there are no windows open, an error occurs. Read-only.
- `Selection As Selection  (read-only)`  
  Returns the Selection object that represents a selected range or the insertion point. Read-only.
- `WordBasic As Object  (read-only)`  
  Returns an automation object (WordBasic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.
- `RecentFiles As RecentFiles  (read-only)`  
  Returns a RecentFiles collection that represents the most recently accessed files.
- `NormalTemplate As Template  (read-only)`  
  Returns a Template object that represents the Normal template.
- `System As System  (read-only)`  
  Returns a System object, which can be used to return system-related information and perform system-related tasks.
- `AutoCorrect As AutoCorrect  (read-only)`  
  Returns an AutoCorrect object that contains the current AutoCorrect options, entries, and exceptions. Read-only.
- `FontNames As FontNames  (read-only)`  
  Returns a FontNames object that includes the names of all the available fonts. Read-only.
- `LandscapeFontNames As FontNames  (read-only)`  
  Returns a FontNames object that includes the names of all the available landscape fonts.
- `PortraitFontNames As FontNames  (read-only)`  
  Returns a FontNames object that includes the names of all the available portrait fonts.
- `Languages As Languages  (read-only)`  
  Returns a Languages collection that represents the proofing languages listed in the Language dialog box.
- `Browser As Browser  (read-only)`  
  Returns a Browser object that represents the Select Browse Object tool on the vertical scroll bar. Read-only.
- `FileConverters As FileConverters  (read-only)`  
  Returns a FileConverters collection that represents all the file converters available to Microsoft Word. Read-only.
- `MailingLabel As MailingLabel  (read-only)`  
  Returns a MailingLabel object that represents a mailing label.
- `Dialogs As Dialogs  (read-only)`  
  Returns a Dialogs collection that represents all the built-in dialog boxes in Word. Read-only.
- `CaptionLabels As CaptionLabels  (read-only)`  
  Returns a CaptionLabels collection that represents all the available caption labels. Read-only.
- `AutoCaptions As AutoCaptions  (read-only)`  
  Returns an AutoCaptions collection that represents the captions that are automatically added when items such as tables and pictures are inserted into a document. Read-only.
- `AddIns As AddIns  (read-only)`  
  Returns an AddIns collection that represents all available add-ins, regardless of whether they're currently loaded. Read-only.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.
- `Version As String  (read-only)`  
  Returns the Microsoft Word version number. Read-only String.
- `ScreenUpdating As Boolean  (read/write)`  
  True if screen updating is turned on. Read/write Boolean.
- `PrintPreview As Boolean  (read/write)`  
  True if print preview is the current view. Read/write Boolean.
- `Tasks As Tasks  (read-only)`  
  Returns a Tasks collection that represents all the applications that are running.
- `SpecialMode As Boolean  (read-only)`  
  True if Microsoft Word is in a special mode (for example, CopyText mode, or MoveText mode). Read-only Boolean.
- `UsableWidth As Long  (read-only)`  
  Returns the maximum width (in points) to which you can set the width of a Microsoft Word document window. Read-only Long.
- `UsableHeight As Long  (read-only)`  
  Returns the maximum height (in points) to which you can set the height of a Microsoft Word document window. Read-only Long.
- `MathCoprocessorAvailable As Boolean  (read-only)`  
  True if a math coprocessor is installed and available to Microsoft Word. Read-only Boolean.
- `MouseAvailable As Boolean  (read-only)`  
  True if there is a mouse available for the system. Read-only Boolean.
- `International As Variant  (read-only)`  
  Returns information about the current country/region and international settings. Read-only Variant.
- `Build As String  (read-only)`  
  Returns the version and build number of the Word application. Read-only String.
- `CapsLock As Boolean  (read-only)`  
  True if the CAPS LOCK key is turned on. Read-only Boolean.
- `NumLock As Boolean  (read-only)`  
  Returns the state of the NUM LOCK key. True if the keys on the numeric keypad insert numbers, False if the keys move the insertion point. Read-only Boolean.
- `UserName As String  (read/write)`  
  Returns or sets the user's name, which is used on envelopes and for the Author document property. Read/write String.
- `UserInitials As String  (read/write)`  
  Returns or sets the user's initials, which Microsoft Word uses to construct comment marks. Read/write String.
- `UserAddress As String  (read/write)`  
  Returns or sets the user's mailing address. Read/write String.
- `MacroContainer As Object  (read-only)`  
  Returns a Template or Document object that represents the template or document in which the module that contains the running procedure is stored.
- `DisplayRecentFiles As Boolean  (read/write)`  
  True if the names of recently used files are displayed on the File menu. Read/write Boolean.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars collection that represents the menu bar and all the toolbars in Microsoft Word.
- `SynonymInfo As SynonymInfo  (read-only)`  
  Returns a SynonymInfo object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the specified word or phrase.
- `VBE As VBE  (read-only)`  
  Returns a VBE object that represents the Visual Basic Editor.
- `DefaultSaveFormat As String  (read/write)`  
  Returns or sets the default format that will appear in the Save as type box in the Save As dialog box. Read/write String.
- `ListGalleries As ListGalleries  (read-only)`  
  Returns a ListGalleries collection that represents the three list template galleries. .
- `ActivePrinter As String  (read/write)`  
  Returns or sets the name of the active printer. Read/write String.
- `Templates As Templates  (read-only)`  
  Returns a Templates collection that represents all the available templates-global templates and those attached to open documents.
- `CustomizationContext As Object  (read/write)`  
  Returns or sets a Template or Document object that represents the template or document in which changes to menu bars, toolbars, and key bindings are stored. Read/write.
- `KeyBindings As KeyBindings  (read-only)`  
  Returns a KeyBindings collection that represents customized key assignments, which include a key code, a key category, and a command.
- `KeysBoundTo As KeysBoundTo  (read-only)`  
  Returns a KeysBoundTo object that represents all the key combinations assigned to the specified item.
- `FindKey As KeyBinding  (read-only)`  
  Returns a KeyBinding object that represents the specified key combination. Read-only.
- `Caption As String  (read/write)`  
  Returns or sets the text displayed in the Title bar of the application window. Read/write String.
- `Path As String  (read-only)`  
  Returns the disk or Web path to the specified object. Read-only String.
- `DisplayScrollBars As Boolean  (read/write)`  
  True if Word displays a scroll bar in at least one document window. False if there are no scroll bars displayed in any window. Read/write Boolean.
- `StartupPath As String  (read/write)`  
  Returns or sets the complete path of the startup folder, excluding the final separator. Read/write String.
- `BackgroundSavingStatus As Long  (read-only)`  
  Returns the number of files queued up to be saved in the background. Read-only Long.
- `BackgroundPrintingStatus As Long  (read-only)`  
  Returns the number of print jobs in the background printing queue. Read-only Long.
- `Left As Long  (read/write)`  
  Returns or sets a Long that represents the horizontal position of the active document, measured in points. Read/write.
- `Top As Long  (read/write)`  
  Returns or sets the vertical position of the active document. Read/write Long.
- `Width As Long  (read/write)`  
  Returns or sets the width of the application window, in points. Read/write Long.
- `Height As Long  (read/write)`  
  Returns or sets the height of the active document window in pixels. Read/write Long.
- `WindowState As WdWindowState  (read/write)`  
  Returns or sets the state of the specified document window or task window. Read/write WdWindowState.
- `DisplayAutoCompleteTips As Boolean  (read/write)`  
  True if Word displays tips that suggest text for completing words, dates, or phrases as you type. Read/write Boolean.
- `Options As Options  (read-only)`  
  Returns an Options object that represents application settings in Microsoft Word.
- `DisplayAlerts As WdAlertLevel  (read/write)`  
  Returns or sets the way certain alerts and messages are handled while a macro is running. Read/write WdAlertLevel.
- `CustomDictionaries As Dictionaries  (read-only)`  
  Returns a Dictionaries object that represents the collection of active custom dictionaries. Read-only.
- `PathSeparator As String  (read-only)`  
  Returns the character used to separate folder names. This property returns a backslash (\). Read-only String.
- `StatusBar As String  (write-only)`  
  This property is no longer supported in Microsoft Word Visual Basic for Applications.
- `MAPIAvailable As Boolean  (read-only)`  
  True if MAPI is installed. Read-only Boolean.
- `DisplayScreenTips As Boolean  (read/write)`  
  True if comments, footnotes, endnotes, and hyperlinks are displayed as tips. Text marked as having comments is highlighted. Read/write Boolean.
- `EnableCancelKey As WdEnableCancelKey  (read/write)`  
  Returns or sets the way that Word handles CTRL+BREAK user interruptions. Read/write WdEnableCancelKey.
- `UserControl As Boolean  (read-only)`  
  True if the document or application was created or opened by the user. Read-only Boolean.
- `MailSystem As WdMailSystem  (read-only)`  
  Returns the mail system (or systems) installed on the host computer. Read-only WdMailSystem.
- `DefaultTableSeparator As String  (read/write)`  
  Returns or sets the single character used to separate text into cells when text is converted to a table. Read/write String.
- `ShowVisualBasicEditor As Boolean  (read/write)`  
  True if the Visual Basic Editor window is visible. Read/write Boolean.
- `BrowseExtraFileTypes As String  (read/write)`  
  Set this property to "text/html" to allow hyperlinked HTML files to be opened in Microsoft Word (instead of the default Internet browser). Read/write String.
- `IsObjectValid As Boolean  (read-only)`  
  True if the specified variable that references an object is valid. Read-only Boolean.
- `HangulHanjaDictionaries As HangulHanjaConversionDictionaries  (read-only)`  
  Returns a HangulHanjaConversionDictionaries collection that represents all the active custom conversion dictionaries.
- `MailMessage As MailMessage  (read-only)`  
  Returns a MailMessage object that represents the active email message.
- `FocusInMailHeader As Boolean  (read-only)`  
  True if the insertion point is in an email header field (the To: field, for example). Read-only Boolean.
- `EmailOptions As EmailOptions  (read-only)`  
  Returns an EmailOptions object that represents the global preferences for email authoring. Read-only.
- `Language As MsoLanguageID  (read-only)`  
  Returns an MsoLanguageID constant that represents the language selected for the Microsoft Word user interface.
- `COMAddIns As COMAddIns  (read-only)`  
  Returns a reference to the COMAddIns collection that represents all the Component Object Model (COM) add-ins currently loaded in Microsoft Word.
- `CheckLanguage As Boolean  (read/write)`  
  True if Microsoft Word automatically detects the language you are using as you type. Read/write Boolean.
- `LanguageSettings As LanguageSettings  (read-only)`  
  Returns a LanguageSettings object, which contains information about the language settings in Microsoft Word.
- `FeatureInstall As MsoFeatureInstall  (read/write)`  
  Returns or sets how Microsoft Word handles calls to methods and properties that require features not yet installed. Read/write MsoFeatureInstall.
- `AutomationSecurity As MsoAutomationSecurity  (read/write)`  
  Returns or sets an MsoAutomationSecurity constant that represents the security setting Microsoft Word uses when programmatically opening files. .
- `FileDialog As FileDialog  (read-only)`  
  Returns a FileDialog object which represents a single instance of a file dialog box.
- `EmailTemplate As String  (read/write)`  
  Returns or sets a String that represents the document template to use when sending email messages. Read/write.
- `NewDocument As NewFile  (read-only)`  
  Returns a NewFile object that represents a document listed on the New tab.
- `ShowStartupDialog As Boolean  (read/write)`  
  True to display the Task Pane when starting Microsoft Word. Read/write Boolean.
- `AutoCorrectEmail As AutoCorrect  (read-only)`  
  Returns an AutoCorrect object that represents automatic corrections made to email messages.
- `TaskPanes As TaskPanes  (read-only)`  
  Returns a TaskPanes collection that represents the most commonly performed tasks in Microsoft Word.
- `DefaultLegalBlackline As Boolean  (read/write)`  
  True for Microsoft Word to compare and merge documents using the Legal blackline option in the Compare and Merge Documents dialog box. Read/write Boolean.
- `XMLNamespaces As XMLNamespaces  (read-only)`  
  Returns an XMLNamespaces collection that represents the XML schemas in the Schema Library.
- `ArbitraryXMLSupportAvailable As Boolean  (read-only)`  
  Returns a Boolean that represents whether Microsoft Word accepts custom XML schemas. True indicates that Word accepts custom XML schemas.
- `Bibliography As Bibliography  (read-only)`  
  Returns a Bibliography object that represents the bibliography references sources stored in Microsoft Word. Read-only.
- `ShowStylePreviews As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word shows a preview of the formatting for styles in the Styles dialog box. Read/write.
- `RestrictLinkedStyles As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word allows linked styles. Read/write.
- `OMathAutoCorrect As OMathAutoCorrect  (read-only)`  
  Returns an OMathAutoCorrect object that represents the auto correct entries for equations. Read-only.
- `DisplayDocumentInformationPanel As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the document properties panel is displayed. Read/write.
- `Assistance As IAssistance  (read-only)`  
  Returns an Assistance object that represents the Microsoft Office Help Viewer. Read-only.
- `OpenAttachmentsInFullScreen As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word opens email attachments in Reading mode. Read/write.
- `ActiveEncryptionSession As Long  (read-only)`  
  Returns a Long that represents the encryption session associated with the active document. Read-only.
- `DontResetInsertionPointProperties As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether Microsoft Word maintains the formatting properties of the text at that position of the Insertion Point after running other code. Read/write.
- `SmartArtLayouts As SmartArtLayouts  (read-only)`  
  Returns a SmartArtLayouts object that represents the set of SmartArt layouts that are currently loaded in the application. Read-only.
- `SmartArtQuickStyles As SmartArtQuickStyles  (read-only)`  
  Returns a SmartArtQuickStyles object that represents the set of SmartArt styles that are currently loaded in the application. Read-only.
- `SmartArtColors As SmartArtColors  (read-only)`  
  Returns a SmartArtColors object that represents the set of color styles that are currently loaded in the application. Read-only.
- `UndoRecord As UndoRecord  (read-only)`  
  Returns an UndoRecord object that provides a custom entry point into the undo stack. Read-only.
- `PickerDialog As PickerDialog  (read-only)`  
  Returns a PickerDialog object that provides the functionality to select people or data in a dialog box. Read-only.
- `ProtectedViewWindows As ProtectedViewWindows  (read-only)`  
  Returns a ProtectedViewWindows collection that represents all Protected View windows. Read-only.
- `ActiveProtectedViewWindow As ProtectedViewWindow  (read-only)`  
  Returns a ProtectedViewWindow object that represents the active Protected View window. Read-only.
- `IsSandboxed As Boolean  (read-only)`  
  True if the application window is a Protected View window. Read-only.
- `FileValidation As MsoFileValidationMode  (read/write)`  
  Returns or sets how Word will validate files before opening them. Read/write MsoFileValidationMode.
- `ChartDataPointTrack As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether charts use cell-reference data-point tracking. Read/write.
- `ShowAnimation As Boolean  (read/write)`  
  This object or member is deprecated and is not intended to be used in your code.
- `SensitivityLabelPolicy As SensitivityLabelPolicy  (read-only)`  
  Returns the SensitivityLabelPolicy object.

## Methods (62)

- `Quit([SaveChanges As Variant], [OriginalFormat As Variant], [RouteDocument As Variant])`  
  Quits Microsoft Word and optionally saves or routes the open documents.
    - `SaveChanges As Variant` (optional): Specifies whether Word saves changed documents before closing. Can be one of the WdSaveOptions constants.
    - `OriginalFormat As Variant` (optional): Specifies the way Word saves documents whose original format was not Word Document format. Can be one of the WdOriginalFormat constants.
    - `RouteDocument As Variant` (optional): True to route the document to the next recipient. If the document does not have a routing slip attached, this argument is ignored.
- `ScreenRefresh()`  
  Updates the display on the monitor with the current information in the video memory buffer.
- `LookupNameProperties(Name As String)`  
  Looks up a name in the global address book list and displays the Properties dialog box, which includes information about the specified name.
    - `Name As String` (required): A name in the global address book.
- `SubstituteFont(UnavailableFont As String, SubstituteFont As String)`  
  Sets font-mapping options.
    - `UnavailableFont As String` (required): The name of a font not available on your computer that you want to map to a different font for display and printing.
    - `SubstituteFont As String` (required): The name of a font available on your computer that you want to substitute for the unavailable font.
- `Repeat([Times As Variant]) As Boolean`  
  Repeats the most recent editing action one or more times. Returns True if the commands were repeated successfully.
    - `Times As Variant` (optional): The number of times you want to repeat the last command.
- `DDEExecute(Channel As Long, Command As String)`  
  Sends a command or series of commands to an application through the specified dynamic data exchange (DDE) channel.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `Command As String` (required): A command or series of commands recognized by the receiving application (the DDE server). If the receiving application cannot perform the specified command, an error occurs.
- `DDEInitiate(App As String, Topic As String) As Long`  
  Opens a dynamic data exchange (DDE) channel to another application, and returns the channel number.
    - `App As String` (required): The name of the application.
    - `Topic As String` (required): The name of a DDE topic&mdash;for example, the name of an open document&mdash;recognized by the application to which you are opening a channel.
- `DDEPoke(Channel As Long, Item As String, Data As String)`  
  Uses an open dynamic data exchange (DDE) channel to send data to an application.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `Item As String` (required): The item within a DDE topic to which the specified data is to be sent.
    - `Data As String` (required): The data to be sent to the receiving application (the DDE server).
- `DDERequest(Channel As Long, Item As String) As String`  
  Uses an open dynamic data exchange (DDE) channel to request information from the receiving application, and returns the information as a String.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `Item As String` (required): The item to be requested.
- `DDETerminate(Channel As Long)`  
  Closes the specified dynamic data exchange (DDE) channel to another application.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
- `DDETerminateAll()`  
  Closes all dynamic data exchange (DDE) channels opened by Microsoft Word.
- `BuildKeyCode(Arg1 As WdKey, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Long`  
  Returns a unique number for the specified key combination.
    - `Arg1 As WdKey` (required): A key you specify by using one of the WdKey constants.
- `KeyString(KeyCode As Long, [KeyCode2 As Variant]) As String`  
  Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A).
    - `KeyCode As Long` (required): A key you specify by using one of the WdKey constants.
    - `KeyCode2 As Variant` (optional): A second key you specify by using one of the WdKey constants.
- `OrganizerCopy(Source As String, Destination As String, Name As String, Object As WdOrganizerObject)`  
  Copies the specified AutoText entry, toolbar, style, or macro project item from the source document or template to the destination document or template.
    - `Source As String` (required): The document or template file name that contains the item you want to copy.
    - `Destination As String` (required): The document or template file name to which you want to copy an item.
    - `Name As String` (required): The name of the AutoText entry, toolbar, style, or macro you want to copy.
    - `Object As WdOrganizerObject` (required): The kind of item you want to copy.
- `OrganizerDelete(Source As String, Name As String, Object As WdOrganizerObject)`  
  Deletes the specified style, AutoText entry, toolbar, or macro project item from a document or template.
    - `Source As String` (required): The file name of the document or template that contains the item you want to delete.
    - `Name As String` (required): The name of the style, AutoText entry, toolbar, or macro you want to delete.
    - `Object As WdOrganizerObject` (required): The kind of item you want to copy.
- `OrganizerRename(Source As String, Name As String, NewName As String, Object As WdOrganizerObject)`  
  Renames the specified style, AutoText entry, toolbar, or macro project item in a document or template.
    - `Source As String` (required): The file name of the document or template that contains the item you want to rename.
    - `Name As String` (required): The name of the style, AutoText entry, toolbar, or macro you want to rename.
    - `NewName As String` (required): The new name for the item.
    - `Object As WdOrganizerObject` (required): The kind of item you want to copy.
- `AddAddress(TagID As SAFEARRAY(String), Value As SAFEARRAY(String))`  
  Adds an entry to the address book. Each entry has values for one or more tag IDs.
    - `TagID As SAFEARRAY(String)` (required): The tag ID values for the new address entry. Each element in the array can contain one of the strings listed in the following table. Only the display name is required; the remaining entries are optional.
    - `Value As SAFEARRAY(String)` (required): The values for the new address entry. Each element corresponds to an element in the TagID array. For more information, see the example.
- `GetAddress([Name As Variant], [AddressProperties As Variant], [UseAutoText As Variant], [DisplaySelectDialog As Variant], [SelectDialog As Variant], [CheckNamesDialog As Variant], [RecentAddressesChoice As Variant], [UpdateRecentAddresses As Variant]) As String`  
  Returns an address from the default address book.
    - `Name As Variant` (optional): The name of the addressee, as specified in the Search Name dialog box in the address book.
    - `AddressProperties As Variant` (optional): If _UseAutoText_ is True, this argument denotes the name of an AutoText entry that defines a sequence of address book properties. If _UseAutoText_ is False or omitted, this argument defines a custom layout. Valid address book property names or sets of property names are surrounded by angle brackets ("<" and ">") and separated by a space or a paragraph mark (for example, "<PR_GIVEN_NAME> <PR_SURNAME>" & vbCr & "<PR_OFFICE_TELEPHONE_NUMBER>"). If the _AddressProperties_ parameter is omitted, a default AutoText entry named "AddressLayout" is used. If "AddressLayout" hasn't been defined, the following address layout definition is used: "<PR_GIVEN_NAME> <PR_SURNAME>" & vbCr & "<PR_STREET_ADDRESS>" & vbCr & "<PR_LOCALITY>" & ", " & "<PR_STATE_OR_PROVINCE>" & " " & "<PR_POSTAL_CODE>" & vbCr & "<PR_COUNTRY>". For a list of the valid address book property names, see the AddAddress method.
    - `UseAutoText As Variant` (optional): True if _AddressProperties_ specifies the name of an AutoText entry that defines a sequence of address book properties; False if it specifies a custom layout.
    - `DisplaySelectDialog As Variant` (optional): Specifies whether the Select Name dialog box is displayed, as shown in the Results table.
    - `SelectDialog As Variant` (optional): Specifies how the Select Name dialog box should be displayed (that is, in what mode), as shown in the Display mode table.
    - `CheckNamesDialog As Variant` (optional): True to display the Check Names dialog box when the value of the _Name_ argument isn't specific enough.
    - `RecentAddressesChoice As Variant` (optional): True to use the list of recently used return addresses.
    - `UpdateRecentAddresses As Variant` (optional): True to add an address to the list of recently used addresses; False to not add the address. If _SelectDialog_ is set to 1 or 2, this argument is ignored.
- `CheckGrammar(String As String) As Boolean`  
  Checks a string for grammatical errors. Returns a Boolean to indicate whether the string contains grammatical errors. True if the string contains no errors.
    - `String As String` (required): The string you want to check for grammatical errors.
- `CheckSpelling(Word As String, [CustomDictionary As Variant], [IgnoreUppercase As Variant], [MainDictionary As Variant], [CustomDictionary2 As Variant], [CustomDictionary3 As Variant], [CustomDictionary4 As Variant], [CustomDictionary5 As Variant], [CustomDictionary6 As Variant], [CustomDictionary7 As Variant], [CustomDictionary8 As Variant], [CustomDictionary9 As Variant], [CustomDictionary10 As Variant]) As Boolean`  
  Checks a string for spelling errors. Returns a Boolean to indicate whether the string contains spelling errors. True if the string has no spelling errors.
    - `Word As String` (required): The text whose spelling is to be checked.
    - `CustomDictionary As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the custom dictionary.
    - `IgnoreUppercase As Variant` (optional): True if capitalization is ignored. If this argument is omitted, the current value of the IgnoreUppercase property is used.
    - `MainDictionary As Variant` (optional): Either an expression that returns a Dictionary object or the file name of the main dictionary.
    - `CustomDictionary2 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary3 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary4 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary5 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary6 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary7 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary8 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary9 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
    - `CustomDictionary10 As Variant` (optional): Either an expression that returns a Dictionary object or the file name of an additional custom dictionary.
- `ResetIgnoreAll()`  
  Clears the list of words that were previously ignored during a spelling check.
- `GetSpellingSuggestions(Word As String, [CustomDictionary As Variant], [IgnoreUppercase As Variant], [MainDictionary As Variant], [SuggestionMode As Variant], [CustomDictionary2 As Variant], [CustomDictionary3 As Variant], [CustomDictionary4 As Variant], [CustomDictionary5 As Variant], [CustomDictionary6 As Variant], [CustomDictionary7 As Variant], [CustomDictionary8 As Variant], [CustomDictionary9 As Variant], [CustomDictionary10 As Variant]) As SpellingSuggestions`  
  Returns a SpellingSuggestions collection that represents the words suggested as spelling replacements for a given word.
    - `Word As String` (required): The word whose spelling is to be checked.
    - `IgnoreUppercase As Variant` (optional): True to ignore words in all uppercase letters. If this argument is omitted, the current value of the IgnoreUppercase property is used.
    - `SuggestionMode As Variant` (optional): Specifies the way Word makes spelling suggestions. Can be one of the following WdSpellingWordType constants: wdAnagram, wdSpellword, or wdWildcard. The default value is WdSpellword.
- `GoBack()`  
  Moves the insertion point among the last three locations where editing occurred in the active document (the same as pressing SHIFT+F5).
- `Help(HelpType As Variant)`  
  Displays installed Help information.
    - `HelpType As Variant` (required): The on-line Help topic or window. Can be any of these WdHelpType constants: wdHelp, wdHelpAbout, wdHelpActiveWindow, wdHelpContents, wdHelpHWP, wdHelpIchitaro, wdHelpIndex, wdHelpPE2, wdHelpPSSHelp, wdHelpSearch, wdHelpUsingHelp. (Some of the constants listed here may not be available to you, depending on the language that you have selected or installed.)
- `HelpTool()`
- `NewWindow() As Window`  
  Opens a new window with the same document as the specified window. Returns a Window object.
- `ListCommands(ListAllCommands As Boolean)`  
  Creates a new document and then inserts a table of Word commands along with their associated shortcut keys and menu assignments.
    - `ListAllCommands As Boolean` (required): True to include all Word commands and their assignments (whether customized or built-in). False to include only commands with customized assignments.
- `ShowClipboard()`  
  Displays the Clipboard task pane.
- `OnTime(When As Variant, Name As String, [Tolerance As Variant])`  
  Starts a background timer that runs a macro at a specified time.
    - `When As Variant` (required): The time at which the macro is to be run.
    - `Name As String` (required): The name of the macro to be run.
    - `Tolerance As Variant` (optional): The maximum time (in seconds) that can elapse before a macro that was not run at the time specified by When is canceled. Macros may not always run at the specified time. For example, if a sort operation is under way or a dialog box is being displayed, the macro will be delayed until Word has completed the task. If this argument is 0 (zero) or omitted, the macro is run regardless of how much time has elapsed since the time specified by When.
- `NextLetter()`
- `CleanString(String As String) As String`  
  Removes nonprinting characters (character codes 1&ndash;29) and special Word characters from the specified string or changes them to spaces (character code 32). Returns the result as a String.
    - `String As String` (required): The source string.
- `ChangeFileOpenDirectory(Path As String)`  
  Sets the folder in which Word searches for documents.
    - `Path As String` (required): The path to the folder in which Word searches for documents.
- `GoForward()`  
  Moves the insertion point forward among the last three locations where editing occurred in the active document.
- `Move(Left As Long, Top As Long)`  
  Positions a task window or the active document window.
    - `Left As Long` (required): The horizontal screen position of the specified window.
    - `Top As Long` (required): The vertical screen position of the specified window.
- `Resize(Width As Long, Height As Long)`  
  Sizes the Word application window or the specified task window.
    - `Width As Long` (required): The width of the window, in points.
    - `Height As Long` (required): The height of the window, in points.
- `InchesToPoints(Inches As Single) As Single`  
  Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a Single.
    - `Inches As Single` (required): The inch value to be converted to points.
- `CentimetersToPoints(Centimeters As Single) As Single`  
  Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a Single.
    - `Centimeters As Single` (required): The centimeter value to be converted to points.
- `MillimetersToPoints(Millimeters As Single) As Single`  
  Converts a measurement from millimeters to points (1 mm = 2.85 points). Returns the converted measurement as a Single.
    - `Millimeters As Single` (required): The millimeter value to be converted to points.
- `PicasToPoints(Picas As Single) As Single`  
  Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a Single.
    - `Picas As Single` (required): The pica value to be converted to points.
- `LinesToPoints(Lines As Single) As Single`  
  Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a Single.
    - `Lines As Single` (required): The line value to be converted to points.
- `PointsToInches(Points As Single) As Single`  
  Converts a measurement from points to inches (1 inch = 72 points). Returns the converted measurement as a Single.
    - `Points As Single` (required): The measurement, in points.
- `PointsToCentimeters(Points As Single) As Single`  
  Converts a measurement from points to centimeters (1 centimeter = 28.35 points). Returns the converted measurement as a Single.
    - `Points As Single` (required): The measurement, in points.
- `PointsToMillimeters(Points As Single) As Single`  
  Converts a measurement from points to millimeters (1 millimeter = 2.835 points). Returns the converted measurement as a Single.
    - `Points As Single` (required): The measurement, in points.
- `PointsToPicas(Points As Single) As Single`  
  Converts a measurement from points to picas (1 pica = 12 points). Returns the converted measurement as a Single.
    - `Points As Single` (required): The measurement, in points.
- `PointsToLines(Points As Single) As Single`  
  Converts a measurement from points to lines (1 line = 12 points). Returns the converted measurement as a Single.
    - `Points As Single` (required): The measurement, in points.
- `Activate()`  
  Activates the specified object.
- `PointsToPixels(Points As Single, [fVertical As Variant]) As Single`  
  Converts a measurement from points to pixels. Returns the converted measurement as a Single.
    - `Points As Single` (required): The point value to be converted to pixels.
    - `fVertical As Variant` (optional): True to return the result as vertical pixels; False to return the result as horizontal pixels.
- `PixelsToPoints(Pixels As Single, [fVertical As Variant]) As Single`  
  Converts a measurement from pixels to points. Returns the converted measurement as a Single.
    - `Pixels As Single` (required): The pixel value to be converted to points.
    - `fVertical As Variant` (optional): True to convert vertical pixels; False to convert horizontal pixels.
- `KeyboardLatin()`  
  Sets the keyboard language to a left-to-right language and the text entry direction to left-to-right.
- `KeyboardBidi()`  
  Sets the keyboard language to a right-to-left language and the text entry direction to right-to-left.
- `ToggleKeyboard()`  
  Switches the keyboard language setting between right-to-left and left-to-right languages.
- `Keyboard([LangId As Long]) As Long`  
  Returns or sets the keyboard language and layout settings.
    - `LangId As Long` (optional): The language and layout combination to which Microsoft Word sets the keyboard. If this argument is omitted, the method returns the current language and layout setting.
- `ProductCode() As String`  
  Returns the Microsoft Word globally unique identifier (GUID) as a String.
- `DefaultWebOptions() As DefaultWebOptions`  
  Returns the DefaultWebOptions object that contains global application-level attributes used by Microsoft Word whenever you save a document as a webpage or open a webpage.
- `SetDefaultTheme(Name As String, DocumentType As WdDocumentMedium)`  
  Sets a default theme for Word to use with new documents, email messages, or webpages.
    - `Name As String` (required): The name of the theme you want to assign as the default theme plus any theme formatting options you want to apply. The format of this string is "themennn" where _theme_ and _nnn_ are defined in the Themes table.
    - `DocumentType As WdDocumentMedium` (required): The type of new document to which you are assigning a default theme.
- `GetDefaultTheme(DocumentType As WdDocumentMedium) As String`  
  Returns a String that represents the name of the default theme plus the theme formatting options Microsoft Word uses for new documents, email messages, or Web pages.
    - `DocumentType As WdDocumentMedium` (required): The type of new document for which you want to retrieve the default theme name.
- `Run(MacroName As String, [varg1 As Variant], [varg2 As Variant], [varg3 As Variant], [varg4 As Variant], [varg5 As Variant], [varg6 As Variant], [varg7 As Variant], [varg8 As Variant], [varg9 As Variant], [varg10 As Variant], [varg11 As Variant], [varg12 As Variant], [varg13 As Variant], [varg14 As Variant], [varg15 As Variant], [varg16 As Variant], [varg17 As Variant], [varg18 As Variant], [varg19 As Variant], [varg20 As Variant], [varg21 As Variant], [varg22 As Variant], [varg23 As Variant], [varg24 As Variant], [varg25 As Variant], [varg26 As Variant], [varg27 As Variant], [varg28 As Variant], [varg29 As Variant], [varg30 As Variant]) As Variant`  
  Runs a Visual Basic macro.
    - `MacroName As String` (required): The name of the macro.
- `PrintOut([Background As Variant], [Append As Variant], [Range As Variant], [OutputFileName As Variant], [From As Variant], [To As Variant], [Item As Variant], [Copies As Variant], [Pages As Variant], [PageType As Variant], [PrintToFile As Variant], [Collate As Variant], [FileName As Variant], [ActivePrinterMacGX As Variant], [ManualDuplexPrint As Variant], [PrintZoomColumn As Variant], [PrintZoomRow As Variant], [PrintZoomPaperWidth As Variant], [PrintZoomPaperHeight As Variant])`  
  Prints all or part of the specified document.
    - `Background As Variant` (optional): Set to True to have the macro continue while Microsoft Word prints the document.
    - `Append As Variant` (optional): Set to True to append the specified document to the file name specified by the OutputFileName argument. False to overwrite the contents of OutputFileName.
    - `Range As Variant` (optional): The page range. Can be any WdPrintOutRange constant.
    - `OutputFileName As Variant` (optional): If PrintToFile is True, this argument specifies the path and file name of the output file.
    - `From As Variant` (optional): The starting page number when Range is set to wdPrintFromTo.
    - `To As Variant` (optional): The ending page number when Range is set to wdPrintFromTo.
    - `Item As Variant` (optional): The item to be printed. Can be any WdPrintOutItem constant.
    - `Copies As Variant` (optional): The number of copies to be printed.
    - `Pages As Variant` (optional): The page numbers and page ranges to be printed, separated by commas. For example, "2, 6-10" prints page 2 and pages 6 through 10.
    - `PageType As Variant` (optional): The type of pages to be printed. Can be any WdPrintOutPages constant.
    - `PrintToFile As Variant` (optional): True to send printer instructions to a file. Make sure to specify a file name with OutputFileName.
    - `Collate As Variant` (optional): When printing multiple copies of a document, True to print all pages of the document before printing the next copy.
    - `FileName As Variant` (optional): The path and file name of the document to be printed. If this argument is omitted, Word prints the active document. (Available only with the Application object.)
    - `ActivePrinterMacGX As Variant` (optional): This argument is available only in Microsoft Office Macintosh Edition. For additional information about this argument, consult the language reference Help included with Microsoft Office Macintosh Edition.
    - `ManualDuplexPrint As Variant` (optional): True to print a two-sided document on a printer without a duplex printing kit. If this argument is True, the PrintBackground and PrintReverse properties are ignored. Use the PrintOddPagesInAscendingOrder and PrintEvenPagesInAscendingOrder properties to control the output during manual duplex printing. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `PrintZoomColumn As Variant` (optional): The number of pages you want Word to fit horizontally on one page. Can be 1, 2, 3, or 4. Use with the PrintZoomRow argument to print multiple pages on a single sheet.
    - `PrintZoomRow As Variant` (optional): The number of pages you want Word to fit vertically on one page. Can be 1, 2, or 4. Use with the PrintZoomColumn argument to print multiple pages on a single sheet.
    - `PrintZoomPaperWidth As Variant` (optional): The width to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).
    - `PrintZoomPaperHeight As Variant` (optional): The height to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).
- `PutFocusInMailHeader()`  
  Places the insertion point in the To line of the mail header if the document in the active window is an email document.
- `LoadMasterList(FileName As String)`  
  Loads a bibliography source file.
    - `FileName As String` (required): The path and file name of the bibliography source file.
- `CompareDocuments(OriginalDocument As Document, RevisedDocument As Document, [Destination As WdCompareDestination], [Granularity As WdGranularity], [CompareFormatting As Boolean], [CompareCaseChanges As Boolean], [CompareWhitespace As Boolean], [CompareTables As Boolean], [CompareHeaders As Boolean], [CompareFootnotes As Boolean], [CompareTextboxes As Boolean], [CompareFields As Boolean], [CompareComments As Boolean], [CompareMoves As Boolean], [RevisedAuthor As String], [IgnoreAllComparisonWarnings As Boolean]) As Document`  
  Compares two documents and returns a Document object that represents the document that contains the differences between the two documents, marked using tracked changes.
    - `OriginalDocument As Document` (required): Specifies the path and file name of the original document.
    - `RevisedDocument As Document` (required): Specifies the path and file name of the revised document to which to compare the original document.
    - `Destination As WdCompareDestination` (optional): Specifies whether to create a new file or whether to mark the differences between the two documents in the original document or in the revised document. Default value is wdCompareDestinationNew.
    - `Granularity As WdGranularity` (optional): Specifies whether changes are tracked by character or by word. Default value is wdGranularityWordLevel.
    - `CompareFormatting As Boolean` (optional): Specifies whether to mark differences in formatting between the two documents. Default value is True.
    - `CompareCaseChanges As Boolean` (optional): Specifies whether to mark differences in case between the two documents. Default value is True.
    - `CompareWhitespace As Boolean` (optional): Specifies whether to mark differences in white space, such as paragraphs or spaces, between the two documents. Default value is True.
    - `CompareTables As Boolean` (optional): Specifies whether to compare the differences in data contained in tables between the two documents. Default value is True.
    - `CompareHeaders As Boolean` (optional): Specifies whether to compare differences in headers and footers between the two documents. Default value is True.
    - `CompareFootnotes As Boolean` (optional): Specifies whether to compare differences in footnotes and endnotes between the two documents. Default value is True.
    - `CompareTextboxes As Boolean` (optional): Specifies whether to compare differences in the data contained within text boxes between the two documents. Default value is True.
    - `CompareFields As Boolean` (optional): Specifies whether to compare differences in fields between the two documents. Default value is True.
    - `CompareComments As Boolean` (optional): Specifies whether to compare differences in comments between the two documents. Default value is True.
    - `CompareMoves As Boolean` (optional): Specifies whether to compare differences in moves between the two documents. Default value is True.
    - `RevisedAuthor As String` (optional): Specifies the name of the person to whom to attribute changes when comparing the two documents.
    - `IgnoreAllComparisonWarnings As Boolean` (optional): Specifies whether to ignore warnings when comparing the two documents.
- `MergeDocuments(OriginalDocument As Document, RevisedDocument As Document, [Destination As WdCompareDestination], [Granularity As WdGranularity], [CompareFormatting As Boolean], [CompareCaseChanges As Boolean], [CompareWhitespace As Boolean], [CompareTables As Boolean], [CompareHeaders As Boolean], [CompareFootnotes As Boolean], [CompareTextboxes As Boolean], [CompareFields As Boolean], [CompareComments As Boolean], [CompareMoves As Boolean], [OriginalAuthor As String], [RevisedAuthor As String], [FormatFrom As WdMergeFormatFrom]) As Document`  
  Compares two documents and returns a Document object that represents the document that contains the differences between the two documents, marked using tracked changes.
    - `OriginalDocument As Document` (required): Specifies the path and file name of the original document.
    - `RevisedDocument As Document` (required): Specifies the path and file name of the revised document to which to compare the original document.
    - `Destination As WdCompareDestination` (optional): Specifies whether to create a new file or whether to mark the differences between the two documents in the original document or in the revised document. Default value is wdCompareDestinationNew.
    - `Granularity As WdGranularity` (optional): Specifies whether changes are tracked by character or by word. Default value is wdGranularityWordLevel.
    - `CompareFormatting As Boolean` (optional): Specifies whether to mark differences in formatting between the two documents. Default value is True.
    - `CompareCaseChanges As Boolean` (optional): Specifies whether to mark differences in case between the two documents. Default value is True.
    - `CompareWhitespace As Boolean` (optional): Specifies whether to mark differences in white space, such as paragraphs or spaces, between the two documents. Default value is True.
    - `CompareTables As Boolean` (optional): Specifies whether to compare the differences in data contained in tables between the two documents. Default value is True.
    - `CompareHeaders As Boolean` (optional): Specifies whether to compare differences in headers and footers between the two documents. Default value is True.
    - `CompareFootnotes As Boolean` (optional): Specifies whether to compare differences in footnotes and endnotes between the two documents. Default value is True.
    - `CompareTextboxes As Boolean` (optional): Specifies whether to compare differences in the data contained within text boxes between the two documents. Default value is True.
    - `CompareFields As Boolean` (optional): Specifies whether to compare differences in fields between the two documents. Default value is True.
    - `CompareComments As Boolean` (optional): Specifies whether to compare differences in comments between the two documents. Default value is True.
    - `OriginalAuthor As String` (optional): Specifies the name of the author of the original document.
    - `RevisedAuthor As String` (optional): Specifies the name of the person to use for unattributed changes after merging two documents.
    - `FormatFrom As WdMergeFormatFrom` (optional): Specifies the document from which to retain formatting.

## Events (34)

- `Quit()`  
  Occurs when the user exits Microsoft Word.
- `DocumentChange()`  
  Occurs when a new document is created, when an existing document is opened, or when another document is made the active document.
- `DocumentOpen(Doc As Document)`  
  Occurs when a document is opened.
    - `Doc As Document` (required): The document that's being opened.
- `DocumentBeforeClose(Doc As Document, Cancel As Boolean)`  
  Occurs immediately before any open document closes.
    - `Doc As Document` (required): The document that's being closed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the document doesn't close when the procedure is finished.
- `DocumentBeforePrint(Doc As Document, Cancel As Boolean)`  
  Occurs before any open document is printed.
    - `Doc As Document` (required): The document that's being printed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the document isn't printed when the procedure is finished.
- `DocumentBeforeSave(Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)`  
  Occurs before any open document is saved.
    - `Doc As Document` (required): The document that is being saved.
    - `SaveAsUI As Boolean` (required): True if the Save As dialog box is displayed, whether to save a new document, in response to the Save command; or in response to the Save As command; or in response to the SaveAs or SaveAs2 method.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the document is not saved when the procedure is finished.
- `NewDocument(Doc As Document)`  
  Occurs when a new document is created.
    - `Doc As Document` (required): The new document.
- `WindowActivate(Doc As Document, Wn As Window)`  
  Occurs when any document window is activated.
    - `Doc As Document` (required): The document displayed in the activated window.
    - `Wn As Window` (required): The window that's being activated.
- `WindowDeactivate(Doc As Document, Wn As Window)`  
  Occurs when any document window is deactivated.
    - `Doc As Document` (required): The document displayed in the deactivated window.
    - `Wn As Window` (required): The deactivated window.
- `WindowSelectionChange(Sel As Selection)`  
  Occurs when the selection changes in the active document window.
    - `Sel As Selection` (required): The text selected. If no text is selected, the Sel parameter returns either nothing or the first character to the right of the insertion point.
- `WindowBeforeRightClick(Sel As Selection, Cancel As Boolean)`  
  Occurs when the editing area of a document window is right-clicked, before the default right-click action.
    - `Sel As Selection` (required): The current selection.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default right-click action does not occur when the procedure is finished.
- `WindowBeforeDoubleClick(Sel As Selection, Cancel As Boolean)`  
  Occurs when the editing area of a document window is double-clicked, before the default double-click action.
    - `Sel As Selection` (required): The current selection.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default double-click action does not occur when the procedure is finished.
- `EPostagePropertyDialog(Doc As Document)`  
  Occurs when a user clicks the E-postage Properties (Labels and Envelopes dialog box) button or Print Electronic Postage button.
    - `Doc As Document` (required): The name of the document to which to add electronic postage.
- `EPostageInsert(Doc As Document)`  
  Occurs when a user inserts electronic postage into a document.
    - `Doc As Document` (required): The name of the document to which to add electronic postage.
- `MailMergeAfterMerge(Doc As Document, DocResult As Document)`  
  Occurs after all records in a mail merge have merged successfully.
    - `Doc As Document` (required): The mail merge main document.
    - `DocResult As Document` (required): The document created from the mail merge
- `MailMergeAfterRecordMerge(Doc As Document)`  
  Occurs after each record in the data source successfully merges in a mail merge.
    - `Doc As Document` (required): The mail merge main document.
- `MailMergeBeforeMerge(Doc As Document, StartRecord As Long, EndRecord As Long, Cancel As Boolean)`  
  Occurs when a merge is executed before any records merge.
    - `Doc As Document` (required): The mail merge main document.
    - `StartRecord As Long` (required): The first record in the data source to include in the mail merge.
    - `EndRecord As Long` (required): The last record in the data source to include in the mail merge.
    - `Cancel As Boolean` (required): True stops the mail merge process before it starts.
- `MailMergeBeforeRecordMerge(Doc As Document, Cancel As Boolean)`  
  Occurs as a merge is executed for the individual records in a merge.
    - `Doc As Document` (required): The mail merge main document.
    - `Cancel As Boolean` (required): True stops the mail merge process, for the current record only, before it starts.
- `MailMergeDataSourceLoad(Doc As Document)`  
  Occurs when the data source is loaded for a mail merge.
    - `Doc As Document` (required): The mail merge main document.
- `MailMergeDataSourceValidate(Doc As Document, Handled As Boolean)`  
  Occurs when a user validates mail merge recipients by clicking Validate in the Mail Merge Recipients dialog box.
    - `Doc As Document` (required): The mail merge main document.
    - `Handled As Boolean` (required): True if the add-in has handled the validation event. This is a forward-only parameter and cannot be set in code. To set this value, you must use the MailMergeDataSourceValidate2 event.
- `MailMergeWizardSendToCustom(Doc As Document)`  
  Occurs when the custom button is clicked during step six of the Mail Merge Wizard.
    - `Doc As Document` (required): The mail merge main document.
- `MailMergeWizardStateChange(Doc As Document, FromState As Long, ToState As Long, Handled As Boolean)`  
  Occurs when a user changes from a specified step to a specified step in the Mail Merge Wizard.
    - `Doc As Document` (required): The mail merge main document.
    - `FromState As Long` (required): The Mail Merge Wizard step from which a user is moving.
    - `ToState As Long` (required): The Mail Merge Wizard step to which a user is moving.
    - `Handled As Boolean` (required): True moves the user to the next step. False for the user to remain at the current step.
- `WindowSize(Doc As Document, Wn As Window)`  
  Occurs when the application window is resized or moved.
    - `Doc As Document` (required): The document in the window being sized.
    - `Wn As Window` (required): The window being sized.
- `XMLSelectionChange(Sel As Selection, OldXMLNode As XMLNode, NewXMLNode As XMLNode, Reason As Long)`  
  Occurs when the parent XML node of the current selection changes.
    - `Sel As Selection` (required): The text selected, including XML elements. If no text is selected, the Sel parameter returns either nothing or the first character to the right of the insertion point.
    - `OldXMLNode As XMLNode` (required): The XML node from which the insertion point is moving.
    - `NewXMLNode As XMLNode` (required): The XML node to which the insertion point is moving.
- `XMLValidationError(XMLNode As XMLNode)`  
  Occurs when there is a validation error in the document.
    - `XMLNode As XMLNode` (required): The XML element that is invalid.
- `DocumentSync(Doc As Document, SyncEventType As MsoSyncEventType)`
    - `Doc As Document` (required): The document being synchronized.
    - `SyncEventType As MsoSyncEventType` (required): The status of the document synchronization.
- `EPostageInsertEx(Doc As Document, cpDeliveryAddrStart As Long, cpDeliveryAddrEnd As Long, cpReturnAddrStart As Long, cpReturnAddrEnd As Long, xaWidth As Long, yaHeight As Long, bstrPrinterName As String, bstrPaperFeed As String, fPrint As Boolean, fCancel As Boolean)`  
  Occurs when a user inserts electronic postage into a document.
    - `Doc As Document` (required): The document to which electronic postage is being added.
    - `cpDeliveryAddrStart As Long` (required): The starting position in the document for the delivery address. Positioning corresponds to the value of the Start property for a Range object.
    - `cpDeliveryAddrEnd As Long` (required): The ending position in the document for the delivery address. Positioning corresponds to the value of the End property for a Range object.
    - `cpReturnAddrStart As Long` (required): The starting position in the document for the return address. Positioning corresponds to the value of the Start property for a Range object.
    - `cpReturnAddrEnd As Long` (required): The ending position in the document for the return address. Positioning corresponds to the value of the End property for a Range object.
    - `xaWidth As Long` (required): The width of the envelope in 1/1440-inch units.
    - `yaHeight As Long` (required): The height of the envelope in 1/1440-inch units.
    - `bstrPrinterName As String` (required): The name of the printer as specified on the Printing Options tab of the Envelope Options dialog box.
    - `bstrPaperFeed As String` (required): The feed method as specified on the Printing Options tab of the Envelope Options dialog box.
    - `fPrint As Boolean` (required): True if the user has specified to print the envelope. False if the user has specified to insert the envelope into the document.
    - `fCancel As Boolean` (required): True cancels inserting the postage.
- `MailMergeDataSourceValidate2(Doc As Document, Handled As Boolean)`  
  Occurs when a user validates mail merge recipients by clicking the Validate addresses link button in the Mail Merge Recipients dialog box.
    - `Doc As Document` (required): The mail merge main document.
    - `Handled As Boolean` (required): True if the add-in has handled the validation event.
- `ProtectedViewWindowOpen(PvWindow As ProtectedViewWindow)`  
  Occurs when a Protected View window is opened.
    - `PvWindow As ProtectedViewWindow` (required): The Protected View window that is opened.
- `ProtectedViewWindowBeforeEdit(PvWindow As ProtectedViewWindow, Cancel As Boolean)`  
  Occurs immediately before editing is enabled on the document in the specified Protected View window.
    - `PvWindow As ProtectedViewWindow` (required): The Protected View window that contains the document that is enabled for editing.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, editing is not enabled on the document.
- `ProtectedViewWindowBeforeClose(PvWindow As ProtectedViewWindow, CloseReason As Long, Cancel As Boolean)`  
  Occurs immediately before a Protected View window or a document in a Protected View window closes.
    - `PvWindow As ProtectedViewWindow` (required): The Protected View window that is closed.
    - `CloseReason As Long` (required): A constant in the WdProtectedViewCloseReason enumeration that specifies the reason the Protected View window is closed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the window does not close when the procedure is finished.
- `ProtectedViewWindowSize(PvWindow As ProtectedViewWindow)`
    - `PvWindow As ProtectedViewWindow` (required): The Protected View window that is sized.
- `ProtectedViewWindowActivate(PvWindow As ProtectedViewWindow)`  
  Occurs when any Protected View window is activated.
    - `PvWindow As ProtectedViewWindow` (required): The Protected View window that is activated.
- `ProtectedViewWindowDeactivate(PvWindow As ProtectedViewWindow)`  
  Occurs when a Protected View window is deactivated.
    - `PvWindow As ProtectedViewWindow` (required): The deactivated Protected View window.
