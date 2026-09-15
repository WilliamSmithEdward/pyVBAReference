# Presentation

**Type:** Class  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493444-5A91-11CF-8700-00AA0060263B}  

Represents a Microsoft PowerPoint presentation.

**Remarks:** The Presentation object is a member of the Presentations collection. The Presentations collection contains all the Presentation objects that represent open presentations in PowerPoint. The following examples describe how to: - Return a presentation that you specify by name or index number - Return the presentation in the active window - Return the presentation in any document window or slide show window you specify

**Example:**

```vba
Presentations("Sample Presentation").Slides.Add 1, 1
```

## Properties (75)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `SlideMaster As _Master  (read-only)`  
  Returns a Master object that represents the slide master.
- `TitleMaster As _Master  (read-only)`  
  Returns a Master object that represents the title master for the specified presentation.
- `HasTitleMaster As MsoTriState  (read-only)`  
  MsoTrue if the specified presentation has a title master. Read-only.
- `TemplateName As String  (read-only)`  
  Returns the name of the first design/master associated with the specified presentation. Read-only.
- `NotesMaster As _Master  (read-only)`  
  Returns a Master object that represents the notes master. Read-only.
- `HandoutMaster As _Master  (read-only)`  
  Returns a Master object that represents the handout master. Read-only.
- `Slides As Slides  (read-only)`  
  Returns a Slides collection that represents all slides in the specified presentation. Read-only.
- `PageSetup As PageSetup  (read-only)`  
  Returns a PageSetup object whose properties control slide setup attributes for the specified presentation. Read-only.
- `ColorSchemes As ColorSchemes  (read-only)`  
  Returns a ColorSchemes collection that represents the color schemes in the specified presentation. Read-only.
- `ExtraColors As ExtraColors  (read-only)`  
  Returns an ExtraColors object that represents the extra colors available in the specified presentation. Read-only.
- `SlideShowSettings As SlideShowSettings  (read-only)`  
  Returns a SlideShowSettings object that represents the slide show settings for the specified presentation. Read-only.
- `Fonts As Fonts  (read-only)`  
  Returns a Fonts collection that represents all fonts used in the specified presentation. Read-only.
- `Windows As DocumentWindows  (read-only)`  
  Returns a DocumentWindows collection that represents all document windows associated with the specified presentation. Read-only.
- `Tags As Tags  (read-only)`  
  Returns a Tags object that represents the tags for the specified object. Read-only.
- `DefaultShape As Shape  (read-only)`  
  Returns a Shape object that represents the default shape for the presentation. Read-only.
- `BuiltInDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the built-in document properties for the specified presentation. Read-only.
- `CustomDocumentProperties As Object  (read-only)`  
  Returns a DocumentProperties collection that represents all the custom document properties for the specified presentation. Read-only.
- `VBProject As VBProject  (read-only)`  
  Returns a VBProject object that represents the individual Visual Basic project for the presentation. Read-only.
- `ReadOnly As MsoTriState  (read-only)`  
  Returns whether the specified presentation is read-only. Read-only.
- `FullName As String  (read-only)`  
  Returns the name of the specified add-in or saved presentation, including the path, the current file system separator, and the file name extension. Read-only String.
- `Name As String  (read-only)`  
  The name of the presentation includes the file name extension (for file types that are registered) but doesn't include its path. You cannot use this property to set the name. Use the SaveAs method to save the presentation under a different name if you need to change the name. Read-only.
- `Path As String  (read-only)`  
  Returns a String that represents the path to the specified Presentation object. Read-only.
- `Saved As MsoTriState  (read/write)`  
  Determines whether changes have been made to a presentation since it was last saved. Read/write.
- `LayoutDirection As PpDirection  (read/write)`  
  Returns or sets the layout direction for the user interface. Read/write.
- `PrintOptions As PrintOptions  (read-only)`  
  Returns a PrintOptions object that represents print options that are saved with the specified presentation. Read-only.
- `Container As Object  (read-only)`  
  Returns the object that contains the specified embedded presentation. Read-only.
- `DisplayComments As MsoTriState  (read/write)`  
  Determines whether comments are displayed in the specified presentation. Read/write.
- `FarEastLineBreakLevel As PpFarEastLineBreakLevel  (read/write)`  
  Returns or sets the line break based upon Asian character level. Read/write.
- `NoLineBreakBefore As String  (read/write)`  
  Returns or sets the characters that cannot begin a line. Read/write.
- `NoLineBreakAfter As String  (read/write)`  
  Returns or sets the characters that cannot end a line. Read/write.
- `SlideShowWindow As SlideShowWindow  (read-only)`  
  Returns a SlideShowWindow object that represents the slide show window in which the specified presentation is running. Read-only.
- `FarEastLineBreakLanguage As MsoFarEastLineBreakLanguageID  (read/write)`  
  Returns or sets the language used to determine which line break level is used when the line break control option is turned on. Read/write.
- `DefaultLanguageID As MsoLanguageID  (read/write)`  
  Returns or sets the default language of presentations. Read/write.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars collection that represents the merged command bar set from the host container application and Microsoft PowerPoint. This property returns a valid object only when the container is a DocObject server, like Microsoft Binder, and PowerPoint is acting as an OLE server. Read-only.
- `EnvelopeVisible As MsoTriState  (read/write)`  
  Determines whether the email message header is visible in the document window. Read/write.
- `VBASigned As MsoTriState  (read-only)`  
  Determines whether the Visual Basic for Applications (VBA) project for the specified document has been digitally signed. Read-only.
- `SnapToGrid As MsoTriState  (read/write)`  
  Determines whether to snap shapes to the gridlines in the specified presentation. Read/write.
- `GridDistance As Single  (read/write)`  
  Sets or returns a Single that represents the distance between gridlines. Read/write.
- `Designs As Designs  (read-only)`  
  Returns a Designs object that represents a collection of designs.
- `Signatures As SignatureSet  (read-only)`  
  Returns a SignatureSet object that represents a collection of digital signatures. Read-only.
- `RemovePersonalInformation As MsoTriState  (read/write)`  
  Determines whether Microsoft PowerPoint should remove all user information from comments and revisions upon saving a presentation. Read/write.
- `PasswordEncryptionProvider As String  (read-only)`  
  Returns the name of the algorithm encryption provider that Microsoft PowerPoint uses when it encrypts documents with passwords. Read-only.
- `PasswordEncryptionAlgorithm As String  (read-only)`  
  Returns the algorithm Microsoft PowerPoint uses for encrypting documents with passwords. Read-only.
- `PasswordEncryptionKeyLength As Long  (read-only)`  
  Returns the key length of the algorithm Microsoft PowerPoint uses when it encrypts documents with passwords. Read-only.
- `PasswordEncryptionFileProperties As Boolean  (read-only)`  
  Returns whether Microsoft PowerPoint encrypts file properties for password-protected documents. Read-only.
- `Password As String  (read/write)`  
  Returns or sets the password that must be supplied to open the specified presentation. Read/write.
- `WritePassword As String  (read/write)`  
  Sets or returns the password for saving changes to the specified document. Read/write.
- `Permission As Permission  (read-only)`
- `SharedWorkspace As SharedWorkspace  (read-only)`
- `Sync As Sync  (read-only)`
- `DocumentLibraryVersions As DocumentLibraryVersions  (read-only)`  
  Returns a DocumentLibraryVersions collection that represents the collection of versions of a shared presentation that has versioning enabled and that is stored in a document library on a server.
- `ContentTypeProperties As MetaProperties  (read-only)`  
  Returns the Microsoft Office MetaProperties collection that describes the metadata stored in the presentation. Read-only.
- `ServerPolicy As ServerPolicy  (read-only)`  
  Returns a Microsoft Office ServerPolicy object. Read-only.
- `DocumentInspectors As DocumentInspectors  (read-only)`  
  Returns the Microsoft Office DocumentInspectors collection. Read-only.
- `HasVBProject As Boolean  (read-only)`  
  Returns whether the active presentation contains a Microsoft Visual Basic for Applications (VBA) project. Read-only.
- `CustomXMLParts As CustomXMLParts  (read-only)`  
  Returns a CustomXMLParts object that represents the collection of custom XML parts associated with the specified Presentation object. Read-only.
- `Final As Boolean  (read/write)`  
  Determines whether the presentation is marked as final (read-only). Read/write.
- `CustomerData As CustomerData  (read-only)`  
  Returns a CustomerData object. Read-only.
- `Research As Research  (read-only)`  
  Returns a Research object that provides access to the research service feature of Microsoft PowerPoint. Read-only.
- `EncryptionProvider As String  (read/write)`  
  Returns a String that specifies the name of the algorithm encryption provider that PowerPoint uses when encrypting documents. Read/write.
- `SectionProperties As SectionProperties  (read-only)`  
  Returns a SectionProperties object. Read-only.
- `Coauthoring As Coauthoring  (read-only)`  
  Returns a Coauthoring object in the current Presentation object. Read-only.
- `Broadcast As Broadcast  (read-only)`  
  Returns the Broadcast object of the current Presentation object. Read-only.
- `HasNotesMaster As Boolean  (read-only)`  
  Indicates whether the presentation has media that resides on a notes master. Read-only.
- `HasHandoutMaster As Boolean  (read-only)`  
  Indicates whether the presentation has media that resides on a handout master. Read-only
- `CreateVideoStatus As PpMediaTaskStatus  (read-only)`  
  Returns the status of creating a video in the current Presentation. Read-only.
- `ChartDataPointTrack As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether charts use cell-reference data-point tracking. Read/write.
- `Guides As Guides  (read-only)`  
  Returns the Guides collection associated with a custom layout. Read-only.
- `WorkIdentity As String  (read/write)`
- `AutoSaveOn As Boolean  (read/write)`  
  True if the edits in the presentation are automatically saved. Read/write Boolean.
- `ReadOnlyRecommended As Boolean  (read-only)`  
  True if the presentation was saved as read-only recommended. Read-only Boolean.
- `SensitivityLabel As ISensitivityLabel  (read-only)`  
  Returns the Microsoft Office SensitivityLabel object from the Presentation.
- `IsFullyDownloaded As Boolean  (read-only)`  
  True if the presentation has finished downloading all of the content. Read-only Boolean.

## Methods (30)

- `AddTitleMaster() As _Master`  
  Adds a title master to the specified presentation and returns a Master object that represents the title master.
- `ApplyTemplate(FileName As String)`  
  Applies a design template to the specified presentation.
    - `FileName As String` (required): Specifies the name of the design template.
- `NewWindow() As DocumentWindow`  
  Opens a new window that contains the specified presentation. Returns a DocumentWindow object that represents the new window.
- `FollowHyperlink(Address As String, [SubAddress As String], [NewWindow As Boolean], [AddHistory As Boolean], [ExtraInfo As String], [Method As MsoExtraInfoMethod], [HeaderInfo As String])`  
  Displays a cached document, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document and displays it in the appropriate application.
    - `Address As String` (required): The address of the target document.
    - `SubAddress As String` (optional): The location in the target document. By default, this argument is an empty string.
    - `NewWindow As Boolean` (optional): True to have the target application opened in a new window. The default value is False.
    - `AddHistory As Boolean` (optional): True to add the link to the current day's history folder.
    - `ExtraInfo As String` (optional): String or byte array that specifies information for HTTP. This argument can be used, for example, to specify the coordinates of an image map or the contents of a form. It can also indicate a FAT file name. The Method argument determines how this extra information is handled.
    - `Method As MsoExtraInfoMethod` (optional): Specifies how ExtraInfo is posted or appended.
    - `HeaderInfo As String` (optional): A string that specifies header information for the HTTP request. The default value is an empty string. You can combine several header lines into a single string by using the following syntax: "string1" & vbCr & "string2". The specified string is automatically converted into ANSI characters. Note that the HeaderInfo argument may overwrite default HTTP header fields.
- `AddToFavorites()`  
  Adds a shortcut that represents the current selection in the specified presentation to the Windows Favorites folder.
- `PrintOut([From As Long], [To As Long], [PrintToFile As String], [Copies As Long], [Collate As MsoTriState])`  
  Prints the specified presentation.
    - `From As Long` (optional): The number of the first page to be printed. If this argument is omitted, printing starts at the beginning of the presentation. Specifying the To and From arguments sets the contents of the PrintRanges object and sets the value of the RangeType property for the presentation.
    - `To As Long` (optional): The number of the last page to be printed. If this argument is omitted, printing continues to the end of the presentation. Specifying the To and From arguments sets the contents of the PrintRanges object and sets the value of the RangeType property for the presentation.
    - `PrintToFile As String` (optional): The name of the file to print to. If you specify this argument, the file is printed to a file rather than sent to a printer. If this argument is omitted, the file is sent to a printer.
    - `Copies As Long` (optional): The number of copies to be printed. If this argument is omitted, only one copy is printed. Specifying this argument sets the value of the NumberOfCopiesproperty.
    - `Collate As MsoTriState` (optional): If this argument is omitted, multiple copies are collated. Specifying this argument sets the value of the Collate property.
- `Save()`  
  Saves the specified presentation.
- `SaveAs(FileName As String, [FileFormat As PpSaveAsFileType], [EmbedTrueTypeFonts As MsoTriState])`  
  Saves a presentation that's never been saved, or saves a previously saved presentation under a different name.
    - `FileName As String` (required): Specifies the name to save the file under. If you don't include a full path, PowerPoint saves the file in the current folder.
    - `FileFormat As PpSaveAsFileType` (optional): Specifies the saved file format. If this argument is omitted, the file is saved in the default file format (ppSaveAsDefault).
- `SaveCopyAs(FileName As String, [FileFormat As PpSaveAsFileType], [EmbedTrueTypeFonts As MsoTriState])`  
  Saves a copy of the specified presentation to a file without modifying the original.
    - `FileName As String` (required): Specifies the name to save the file under. If you don't include a full path, PowerPoint saves the file in the current folder.
    - `FileFormat As PpSaveAsFileType` (optional): The file format.
    - `EmbedTrueTypeFonts As MsoTriState` (optional): Specifies whether TrueType fonts are embedded.
- `Export(Path As String, FilterName As String, [ScaleWidth As Long], [ScaleHeight As Long])`  
  Exports each slide in the presentation, using the specified graphics filter, and saves the exported files in the specified folder.
    - `Path As String` (required): The path of the folder where you want to save the exported slides. You can include a full path; if you don't do this, Microsoft PowerPoint creates a subfolder in the current folder for the exported slides.
    - `FilterName As String` (required): The graphics format in which you want to export slides. The specified graphics format must have an export filter registered in the Windows registry. You can specify either the registered extension or the registered filter name. PowerPoint will first search for a matching extension in the registry. If no extension that matches the specified string is found, PowerPoint will look for a filter name that matches.
    - `ScaleWidth As Long` (optional): The width in pixels of an exported slide.
    - `ScaleHeight As Long` (optional): The height in pixels of an exported slide.
- `Close()`  
  Closes the specified presentation.
- `UpdateLinks()`  
  Updates linked OLE objects in the specified presentation.
- `CheckIn([SaveChanges As Boolean], [Comments As Variant], [MakePublic As Variant])`  
  Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
    - `SaveChanges As Boolean` (optional): True saves the presentation to the server location. The default value is False.
    - `Comments As Variant` (optional): Comments for the revision of the presentation being checked in (only applies if SaveChanges equals True).
    - `MakePublic As Variant` (optional): True allows the user to perform a publish on the presentation after being checked in. This submits the document for the approval process, which can eventually result in a version of the presentation being published to users with read-only rights to the presentation (only applies if SaveChanges equals True).
- `CanCheckIn() As Boolean`  
  Returns True if Microsoft PowerPoint can check in a specified presentation to a server.
- `SetPasswordEncryptionOptions(PasswordEncryptionProvider As String, PasswordEncryptionAlgorithm As String, PasswordEncryptionKeyLength As Long, PasswordEncryptionFileProperties As Boolean)`  
  Sets the options Microsoft PowerPoint uses for encrypting presentations with passwords.
    - `PasswordEncryptionProvider As String` (required): The name of the encryption provider.
    - `PasswordEncryptionAlgorithm As String` (required): The name of the encryption algorithm. PowerPoint supports stream-encrypted algorithms.
    - `PasswordEncryptionKeyLength As Long` (required): The encryption key length. Must be a multiple of 8, starting at 40.
    - `PasswordEncryptionFileProperties As Boolean` (required): msoTrue for PowerPoint to encrypt file properties.
- `SendFaxOverInternet([Recipients As String], [Subject As String], [ShowMessage As Boolean])`  
  Sends a presentation as a fax to the specified recipients.
    - `Recipients As String` (optional): A String that represents the fax numbers and email addresses of the people to whom to send the fax. Separate multiple recipients with a semicolon.
    - `Subject As String` (optional): A String that represents the subject line for the faxed presentation.
    - `ShowMessage As Boolean` (optional): Whether to display the fax message before sending it. True displays the fax message before sending it. False sends the fax without displaying the fax message.
- `RemoveDocumentInformation(Type As PpRemoveDocInfoType)`  
  Removes document information, such as personal information, comments, and document properties, from a Microsoft PowerPoint presentation.
    - `Type As PpRemoveDocInfoType` (required): Type of information to be removed.
- `CheckInWithVersion([SaveChanges As Boolean], [Comments As Variant], [MakePublic As Variant], [VersionType As Variant])`  
  Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
    - `SaveChanges As Boolean` (optional): True saves the presentation to the server location. The default value is False.
    - `Comments As Variant` (optional): Comments for the revision of the presentation being checked in (only applies if SaveChanges equals True).
    - `MakePublic As Variant` (optional): True allows the user to perform a publish on the presentation after being checked in. This submits the document for the approval process, which can eventually result in a version of the presentation being published to users with read-only rights to the presentation (only applies if SaveChanges equals True).
    - `VersionType As Variant` (optional): Version number of the presentation.
- `ExportAsFixedFormat(Path As String, FixedFormatType As PpFixedFormatType, [Intent As PpFixedFormatIntent], [FrameSlides As MsoTriState], [HandoutOrder As PpPrintHandoutOrder], [OutputType As PpPrintOutputType], [PrintHiddenSlides As MsoTriState], [PrintRange As PrintRange], [RangeType As PpPrintRangeType], [SlideShowName As String], [IncludeDocProperties As Boolean], [KeepIRMSettings As Boolean], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [ExternalExporter As Variant])`  
  Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.
    - `Path As String` (required): The path for the export.
    - `FixedFormatType As PpFixedFormatType` (required): The format to which the slides should be exported.
    - `Intent As PpFixedFormatIntent` (optional): The purpose of the export.
    - `FrameSlides As MsoTriState` (optional): Whether the slides to be exported should be bordered by a frame.
    - `HandoutOrder As PpPrintHandoutOrder` (optional): The order in which the handout should be printed.
    - `OutputType As PpPrintOutputType` (optional): The type of output.
    - `PrintHiddenSlides As MsoTriState` (optional): Whether to print hidden slides.
    - `PrintRange As PrintRange` (optional): The slide range, can be Nothing.
    - `RangeType As PpPrintRangeType` (optional): The type of slide range.
    - `SlideShowName As String` (optional): The name of the slide show.
    - `IncludeDocProperties As Boolean` (optional): Whether the document properties should also be exported. The default is False.
    - `KeepIRMSettings As Boolean` (optional): Whether the IRM settings should also be exported.</br></br>If FixedFormatType is PpFixedFormatTypePDF, this flag determines if labels and IRM settings should be exported.</br></br>The default is True.
    - `DocStructureTags As Boolean` (optional): Whether to include document structure tags to improve document accessibility. The default is True.
    - `BitmapMissingFonts As Boolean` (optional): Whether to include a bitmap of the text. The default is True.
    - `ExternalExporter As Variant` (optional): A pointer to an Office add-in that implements the IMsoDocExporter COM interface and allows calls to an alternate implementation of code. The default is a null pointer.
- `GetWorkflowTasks() As WorkflowTasks`  
  Returns the Microsoft Office WorkflowTasks collection.
- `GetWorkflowTemplates() As WorkflowTemplates`  
  Returns the Microsoft Office WorkflowTemplates collection.
- `LockServerFile()`  
  Locks the presentation on the Microsoft Office SharePoint server to prevent its modification.
- `ApplyTheme(themeName As String)`  
  Applies a theme or design template to the specified presentation.
    - `themeName As String` (required): The path and name of the theme file (.thmx) or design template file (.pot) to apply to the Presentation object.
- `EnsureAllMediaUpgraded()`  
  Ensures that all media is up to date in a Presentation object.
- `Convert2(FileName As String)`  
  Converts a file to a different file type.
    - `FileName As String` (required): The name of the file to be converted.
- `CreateVideo(FileName As String, [UseTimingsAndNarrations As Boolean], [DefaultSlideDuration As Long], [VertResolution As Long], [FramesPerSecond As Long], [Quality As Long])`  
  Creates a video in a Presentation object.
    - `FileName As String` (required): The name of the video file to create.
    - `UseTimingsAndNarrations As Boolean` (optional): Indicates whether to use timings and narrations.
    - `DefaultSlideDuration As Long` (optional): The duration, in seconds, to view the slide.
    - `VertResolution As Long` (optional): The resolution of the slide.
    - `FramesPerSecond As Long` (optional): The number of frames per second.
    - `Quality As Long` (optional): The level of quality of the slide.
- `ApplyTemplate2(FileName As String, VariantGUID As String)`  
  Applies a design template and theme variant to the presentation.
- `ExportAsFixedFormat2(Path As String, FixedFormatType As PpFixedFormatType, [Intent As PpFixedFormatIntent], [FrameSlides As MsoTriState], [HandoutOrder As PpPrintHandoutOrder], [OutputType As PpPrintOutputType], [PrintHiddenSlides As MsoTriState], [PrintRange As PrintRange], [RangeType As PpPrintRangeType], [SlideShowName As String], [IncludeDocProperties As Boolean], [KeepIRMSettings As Boolean], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [IncludeMarkup As Boolean], [ExternalExporter As Variant])`  
  Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.
    - `Path As String` (required): The path for the export.
    - `FixedFormatType As PpFixedFormatType` (required): The format to which the slides should be exported.
    - `Intent As PpFixedFormatIntent` (optional): The purpose of the export.
    - `FrameSlides As MsoTriState` (optional): Whether the slides to be exported should be bordered by a frame.
    - `HandoutOrder As PpPrintHandoutOrder` (optional): The order in which the handout should be printed.
    - `OutputType As PpPrintOutputType` (optional): The type of output.
    - `PrintHiddenSlides As MsoTriState` (optional): Whether to print hidden slides.
    - `PrintRange As PrintRange` (optional): The slide range. Can be set to Nothing for all
    - `RangeType As PpPrintRangeType` (optional): The type of slide range.
    - `SlideShowName As String` (optional): The name of the slide show.
    - `IncludeDocProperties As Boolean` (optional): Whether the document properties should also be exported. The default is False.
    - `KeepIRMSettings As Boolean` (optional): Whether the IRM settings should also be exported.</br></br>If FixedFormatType is PpFixedFormatTypePDF, this flag determines if labels and IRM settings should be exported.</br></br>The default is True.
    - `DocStructureTags As Boolean` (optional): Whether to include document structure tags to improve document accessibility. The default is True.
    - `BitmapMissingFonts As Boolean` (optional): Whether to include a bitmap of the text. The default is True.
    - `UseISO19005_1 As Boolean` (optional): Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is False.
    - `IncludeMarkup As Boolean` (optional): Whether the resulting document should include associated pen marks.
    - `ExternalExporter As Variant` (optional): A pointer to an Office add-in that implements the IMsoDocExporter COM interface and allows calls to an alternate implementation of code. The default is a null pointer.
- `SaveCopyAs2(FileName As String, [FileFormat As PpSaveAsFileType], [EmbedTrueTypeFonts As MsoTriState], [ReadOnlyRecommended As MsoTriState])`  
  Saves a copy of the specified presentation to a file without modifying the original.
    - `FileName As String` (required): Specifies the name to save the file under. If you don't include a full path, PowerPoint saves the file in the current folder.
    - `FileFormat As PpSaveAsFileType` (optional): The file format.
    - `EmbedTrueTypeFonts As MsoTriState` (optional): Specifies whether TrueType fonts are embedded.
    - `ReadOnlyRecommended As MsoTriState` (optional): Specifies whether the file should be marked as ReadOnlyRecommended.
- `ExportAsFixedFormat3(Path As String, FixedFormatType As PpFixedFormatType, [Intent As PpFixedFormatIntent], [FrameSlides As MsoTriState], [HandoutOrder As PpPrintHandoutOrder], [OutputType As PpPrintOutputType], [PrintHiddenSlides As MsoTriState], [PrintRange As PrintRange], [RangeType As PpPrintRangeType], [SlideShowName As String], [IncludeDocProperties As Boolean], [KeepIRMSettings As Boolean], [DocStructureTags As Boolean], [BitmapMissingFonts As Boolean], [UseISO19005_1 As Boolean], [IncludeMarkup As Boolean], [Bookmarks As Boolean], [DocumentMarkup As Boolean], [PromotedHyperlinkShape As Boolean], [ExternalExporter As Variant])`  
  Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.
    - `Path As String` (required): The path for the export.
    - `FixedFormatType As PpFixedFormatType` (required): The format to which the slides should be exported.
    - `Intent As PpFixedFormatIntent` (optional): The purpose of the export.
    - `FrameSlides As MsoTriState` (optional): Whether the slides to be exported should be bordered by a frame.
    - `HandoutOrder As PpPrintHandoutOrder` (optional): The order in which the handout should be printed.
    - `OutputType As PpPrintOutputType` (optional): The type of output.
    - `PrintHiddenSlides As MsoTriState` (optional): Whether to print hidden slides.
    - `PrintRange As PrintRange` (optional): The slide range. Can be set to Nothing for all
    - `RangeType As PpPrintRangeType` (optional): The type of slide range.
    - `SlideShowName As String` (optional): The name of the slide show.
    - `IncludeDocProperties As Boolean` (optional): Whether the document properties should also be exported. The default is False.
    - `KeepIRMSettings As Boolean` (optional): Whether the IRM settings should also be exported.</br></br>If FixedFormatType is PpFixedFormatTypePDF, this flag determines if labels and IRM settings should be exported.</br></br>The default is True.
    - `DocStructureTags As Boolean` (optional): Whether to include document structure tags to improve document accessibility. The default is True.
    - `BitmapMissingFonts As Boolean` (optional): Whether to include a bitmap of the text. The default is True.
    - `UseISO19005_1 As Boolean` (optional): Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is False.
    - `IncludeMarkup As Boolean` (optional): Whether the resulting document should include associated pen marks.
    - `Bookmarks As Boolean` (optional): Whether bookmarks for each section and slide should be included in the exported document. When using this option, external exporters should not add their own bookmarks for sections or slides. The default is True.
    - `DocumentMarkup As Boolean` (optional): Whether the Document tag should be included in the document structure tags. When using this option, external exporters should not add their own Document tag. The default is True.
    - `PromotedHyperlinkShape As Boolean` (optional): Whether hyperlinks should be promoted to siblings of objects rather than nested within objects in document structure tags. Transparent text elements with alpha of 0 are included for the hyperlinks and external exporters should respect the alpha value so that they are not visible in the document. The default is True.
    - `ExternalExporter As Variant` (optional): A pointer to an Office add-in that implements the IMsoDocExporter COM interface and allows calls to an alternate implementation of code. The default is a null pointer.
