# Application

**Type:** Class  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493441-5A91-11CF-8700-00AA0060263B}  

Represents the entire Microsoft PowerPoint application.

**Remarks:** The Application object contains: - Application-wide settings and options (the name of the active printer, for example). - Properties that return top-level objects, such as ActivePresentation, and Windows. When you are writing code that will run from PowerPoint, you can use the following properties of the Application object without the object qualifier: ActivePresentation, ActiveWindow, AddIns, Presentations, SlideShowWindows, Windows. For example, instead of writing Application.ActiveWindow.Height = 200, you can write ActiveWindow.Height = 200.

**Example:**

```vba
Dim MyPath As String

MyPath = Application.Path
```

## Properties (50)

- `Presentations As Presentations  (read-only)`  
  Returns a Presentations collection that represents all open presentations. Read-only.
- `Windows As DocumentWindows  (read-only)`  
  Returns a DocumentWindows collection that represents all open document windows. Read-only.
- `ActiveWindow As DocumentWindow  (read-only)`  
  Returns a DocumentWindow object that represents the active document window. Read-only.
- `ActivePresentation As Presentation  (read-only)`  
  Returns a Presentation object that represents the presentation open in the active window. Read-only.
- `SlideShowWindows As SlideShowWindows  (read-only)`  
  Returns a SlideShowWindows collection that represents all open slide show windows. Read-only.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars collection that represents all the command bars in Microsoft PowerPoint. Read-only.
- `Path As String  (read-only)`  
  Returns a String that represents the path to the specified Application object. Read-only.
- `Name As String  (read-only)`  
  Returns the string "Microsoft PowerPoint." Read-only.
- `Caption As String  (read/write)`  
  Returns the text that appears in the title bar of the application window. Read/write.
- `Build As String  (read-only)`  
  Returns the build number for the current instance of Microsoft PowerPoint. Read-only.
- `Version As String  (read-only)`  
  Returns the Microsoft PowerPoint version number. Read-only.
- `OperatingSystem As String  (read-only)`  
  Returns the name of the operating system. Read-only.
- `ActivePrinter As String  (read-only)`  
  Returns the name of the active printer. Read-only.
- `Creator As Long  (read-only)`  
  Returns a Long that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.
- `AddIns As AddIns  (read-only)`  
  Returns the program-specific AddIns collection that represents all the add-ins listed in the Add-Ins dialog box (click the Office button, click PowerPoint Options, click Add-Ins, click PowerPoint Add-Ins on the Manage list). Read-only.
- `VBE As VBE  (read-only)`  
  Returns a VBE object that represents the Visual Basic Editor. Read-only.
- `Left As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the left edge of the document, application, and slide show windows to the left edge of the application window's client area. Setting this property to a very large positive or negative value may position the window completely off the desktop. Read/write.
- `Top As Single  (read/write)`  
  Returns or sets a Single that represents the distance in points from the top edge of the document, application, and slide show window to the top edge of the application window's client area. Read/write.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified object, in points. Read/write.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified object, in points. Read/write.
- `WindowState As PpWindowState  (read/write)`  
  Returns or sets the state of the specified window. Read/write.
- `Visible As MsoTriState  (read/write)`  
  Returns or sets a MsoTriState value that determines whether the object is visible. Read/write.
- `Active As MsoTriState  (read-only)`  
  Returns whether the specified pane or window is active. Read-only.
- `COMAddIns As COMAddIns  (read-only)`  
  Returns a reference to the Component Object Model (COM) add-ins currently loaded in Microsoft PowerPoint. These add-ins are listed on the Add-Ins tab in the PowerPoint Options dialog box. Read-only.
- `ProductCode As String  (read-only)`  
  Returns the Microsoft PowerPoint globally unique identifier (GUID). Read-only.
- `LanguageSettings As LanguageSettings  (read-only)`  
  Returns a LanguageSettings object that contains information about the language settings in Microsoft PowerPoint. Read-only.
- `ShowWindowsInTaskbar As MsoTriState  (read/write)`  
  Determines whether there is a separate Windows taskbar button for each open presentation. Read/write.
- `FeatureInstall As MsoFeatureInstall  (read/write)`  
  Returns or sets how Microsoft PowerPoint handles calls to methods and properties that require features not yet installed. Read/write.
- `FileDialog As FileDialog  (read-only)`  
  Returns a FileDialog object that represents a single instance of a file dialog box. Read-only.
- `DisplayGridLines As MsoTriState  (read/write)`  
  Determines whether to display gridlines in Microsoft PowerPoint. Read/write.
- `AutomationSecurity As MsoAutomationSecurity  (read/write)`  
  Represents the security mode that Microsoft PowerPoint uses when it opens files programmatically. Read/write.
- `NewPresentation As NewFile  (read-only)`  
  Returns a NewFile object that represents a presentation listed on the New Presentation task pane. Read-only.
- `DisplayAlerts As PpAlertLevel  (read/write)`  
  Sets or returns whether Microsoft PowerPoint displays alerts while running a macro. Read/write.
- `ShowStartupDialog As MsoTriState  (read/write)`  
  Determines whether to display the New Presentation task pane when Microsoft PowerPoint is started. Read/write.
- `AutoCorrect As AutoCorrect  (read-only)`  
  Returns an AutoCorrect object that represents the AutoCorrect functionality in Microsoft PowerPoint.
- `Options As Options  (read-only)`  
  Returns an Options object that represents application options in Microsoft PowerPoint.
- `DisplayDocumentInformationPanel As Boolean  (read/write)`  
  Returns or sets whether the Document Properties panel is displayed in the Microsoft PowerPoint user interface. Read/write.
- `Assistance As IAssistance  (read-only)`  
  Gets a reference to the Microsoft Office IAssistance object, which provides a means for developers to create a customized help experience for users within Microsoft Office. Read-only.
- `ActiveEncryptionSession As Long  (read-only)`  
  Represents the encryption session associated with the active presentation. Read-only.
- `FileConverters As FileConverters  (read-only)`  
  Returns information about installed file converters. Returns null if there are no converters installed. Read-only Variant.
- `SmartArtLayouts As SmartArtLayouts  (read-only)`  
  Returns the SmartArt layout of the current Application object. Read-only.
- `SmartArtQuickStyles As SmartArtQuickStyles  (read-only)`  
  Returns the quick styles of the SmartArt diagram in the current Application object. Read-only.
- `SmartArtColors As SmartArtColors  (read-only)`  
  Returns the SmartArt colors of the current Application object. Read-only.
- `ProtectedViewWindows As ProtectedViewWindows  (read-only)`  
  Returns a ProtectedViewWindows collection that represents all the Protected View windows that are open in the application. Read-only
- `ActiveProtectedViewWindow As ProtectedViewWindow  (read-only)`  
  Returns a ProtectedViewWindow object that represents the active Protected View window (the window on top). Read-only.
- `IsSandboxed As Boolean  (read-only)`  
  Returns True if the specified presentation is open in a Protected View window. Read-only.
- `FileValidation As MsoFileValidationMode  (read/write)`  
  Returns or sets a value that indicates how PowerPoint will validate files before opening them. Read/write
- `ChartDataPointTrack As Boolean  (read/write)`  
  Returns or sets a Boolean that specifies whether charts use cell-reference data-point tracking. Read/write.
- `DisplayGuides As MsoTriState  (read/write)`  
  Gets or sets whether drawing guides are displayed in the application.
- `SensitivityLabelPolicy As SensitivityLabelPolicy  (read-only)`  
  Returns the SensitivityLabelPolicy object.

## Methods (6)

- `Help([HelpFile As String], [ContextID As Long])`  
  Displays a Help topic.
    - `HelpFile As String` (optional): The name of the Help file you want to display. Can be either a .chm or an .hlp file. If this argument is not specified, Microsoft PowerPoint Help is used.
    - `ContextID As Long` (optional): The context ID number for the Help topic. If this argument is not specified or if it specifies a context ID number that is not associated with a Help topic, the Help Topics dialog box is displayed.
- `Quit()`  
  Quits Microsoft PowerPoint. This is equivalent to clicking the Office button and then clicking Exit PowerPoint.
- `Run(MacroName As String, safeArrayOfParams As SAFEARRAY(Variant)) As Variant`  
  Runs a Visual Basic procedure.
    - `MacroName As String` (required): The name of the procedure to be run. The string can contain the following: a loaded presentation or add-in file name followed by an exclamation point (!), a valid module name followed by a period (.), and the procedure name. For example, the following is a valid MacroName value: "MyPres.pptm!Module1.Test."
- `Activate()`  
  Activates the specified object.
- `StartNewUndoEntry()`  
  Starts a new undo entry in the Application object.
- `OpenThemeFile(themeFileName As String) As Theme`  
  Opens the specified theme file (*thmx).

## Events (33)

- `WindowSelectionChange(Sel As Selection)`  
  Occurs when the selection of text, a shape, or a slide in the active document window changes, whether in the user interface or in code.
    - `Sel As Selection` (required): Represents the object selected.
- `WindowBeforeRightClick(Sel As Selection, Cancel As Boolean)`  
  Occurs when you right-click a shape, a slide, a notes page, or some text. This event is triggered by the MouseUp event.
    - `Sel As Selection` (required): The selection below the mouse pointer when the right-click occurred.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default context menu does not appear when the procedure is finished.
- `WindowBeforeDoubleClick(Sel As Selection, Cancel As Boolean)`  
  Occurs when you double-click the items in the views listed in the following table.
    - `Sel As Selection` (required): The selection below the mouse pointer when the double-click occurs.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default double-click action isn't performed when the procedure is finished.
- `PresentationClose(Pres As Presentation)`  
  Occurs immediately before any open presentation closes, as it is removed from the Presentations collection.
    - `Pres As Presentation` (required): The presentation that is being closed.
- `PresentationSave(Pres As Presentation)`  
  Occurs before any open presentation is saved.
    - `Pres As Presentation` (required): The presentation to be saved.
- `PresentationOpen(Pres As Presentation)`  
  Occurs after an existing presentation is opened, as it is added to the Presentations collection.
    - `Pres As Presentation` (required): The presentation that is opened.
- `NewPresentation(Pres As Presentation)`  
  Occurs after a presentation is created, as it is added to the Presentations collection.
    - `Pres As Presentation` (required): The new presentation.
- `PresentationNewSlide(Sld As Slide)`  
  Occurs when a new slide is created in any open presentation, as the slide is added to the Slides collection.
    - `Sld As Slide` (required): The new slide.
- `WindowActivate(Pres As Presentation, Wn As DocumentWindow)`  
  Occurs when the application window or any document window is activated.
    - `Pres As Presentation` (required): The presentation displayed in the activated window.
    - `Wn As DocumentWindow` (required): The activated document window.
- `WindowDeactivate(Pres As Presentation, Wn As DocumentWindow)`  
  Occurs when the application window or any document window is deactivated.
    - `Pres As Presentation` (required): The presentation displayed in the deactivated window.
    - `Wn As DocumentWindow` (required): The deactivated document window.
- `SlideShowBegin(Wn As SlideShowWindow)`  
  Occurs when you start a slide show.
    - `Wn As SlideShowWindow` (required): The slide show window initialized prior to this event.
- `SlideShowNextBuild(Wn As SlideShowWindow)`  
  Occurs upon mouse-click or timing animation, but before the animated object becomes visible. .
    - `Wn As SlideShowWindow` (required): The active slide show window.
- `SlideShowNextSlide(Wn As SlideShowWindow)`  
  Occurs immediately before the transition to the next slide. For the first slide, occurs immediately after the SlideShowBegin event.
    - `Wn As SlideShowWindow` (required): The active slide show window.
- `SlideShowEnd(Pres As Presentation)`  
  Occurs after a slide show ends, immediately after the last SlideShowNextSlide event occurs.
    - `Pres As Presentation` (required): The presentation closed when this event occurs.
- `PresentationPrint(Pres As Presentation)`  
  Occurs before a presentation is printed.
    - `Pres As Presentation` (required): The presentation to be printed.
- `SlideSelectionChanged(SldRange As SlideRange)`  
  Occurs at different times depending on the current view.
    - `SldRange As SlideRange` (required): The selection of slides. In most cases this would be a single slide (for example, in Slide View you navigate to the next slide), but in some cases this could be multiple slides (for example, a marquee selection in Slide Sorter View).
- `ColorSchemeChanged(SldRange As SlideRange)`  
  Occurs after a color scheme is changed.
    - `SldRange As SlideRange` (required): The range of slides affected by the change.
- `PresentationBeforeSave(Pres As Presentation, Cancel As Boolean)`  
  Occurs before a presentation is saved.
    - `Pres As Presentation` (required): The presentation being saved.
    - `Cancel As Boolean` (required): True to cancel the save process.
- `SlideShowNextClick(Wn As SlideShowWindow, nEffect As Effect)`  
  Occurs on the next click of the slide.
    - `Wn As SlideShowWindow` (required): The slide show window initialized prior to this event.
    - `nEffect As Effect` (required): The effect to animate on next click.
- `AfterNewPresentation(Pres As Presentation)`  
  Occurs after a presentation is created.
    - `Pres As Presentation` (required): Name of the presentation.
- `AfterPresentationOpen(Pres As Presentation)`  
  Occurs after an existing presentation is opened.
    - `Pres As Presentation` (required): The presentation that is opened.
- `PresentationSync(Pres As Presentation, SyncEventType As MsoSyncEventType)`  
  Occurs when the local copy of a presentation that is part of a Document Workspace is synchronized with the copy on the server. Provides important status information regarding the success or failure of the synchronization of the presentation.
    - `Pres As Presentation` (required): The presentation that is being synchronized.
    - `SyncEventType As MsoSyncEventType` (required): The status of the synchronization.
- `SlideShowOnNext(Wn As SlideShowWindow)`  
  Occurs when the user clicks Next to move within the current slide.
    - `Wn As SlideShowWindow` (required): The active slideshow window.
- `SlideShowOnPrevious(Wn As SlideShowWindow)`  
  Occurs when the user clicks Previous to move within the current slide.
    - `Wn As SlideShowWindow` (required): The active slideshow window.
- `PresentationBeforeClose(Pres As Presentation, Cancel As Boolean)`  
  Represents a Presentation object before it closes.
    - `Pres As Presentation` (required): The Presentation object.
    - `Cancel As Boolean` (required): If set to True, the presentation will not close.
- `ProtectedViewWindowOpen(ProtViewWindow As ProtectedViewWindow)`  
  Occurs when a Protected View window is opened.
    - `ProtViewWindow As ProtectedViewWindow` (required): The Protected View window that is opened.
- `ProtectedViewWindowBeforeEdit(ProtViewWindow As ProtectedViewWindow, Cancel As Boolean)`  
  Occurs immediately before editing is enabled on the document in the specified Protected View window.
    - `ProtViewWindow As ProtectedViewWindow` (required): The Protected View window that contains the document that is enabled for editing.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, editing is not enabled on the document.
- `ProtectedViewWindowBeforeClose(ProtViewWindow As ProtectedViewWindow, ProtectedViewCloseReason As PpProtectedViewCloseReason, Cancel As Boolean)`  
  Occurs immediately before a Protected View window or a document in a Protected View window closes.
    - `ProtViewWindow As ProtectedViewWindow` (required): The Protected View window that is closed.
    - `ProtectedViewCloseReason As PpProtectedViewCloseReason` (required): A constant that specifies the reason the Protected View window is closed.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the window does not close when the procedure is finished.
- `ProtectedViewWindowActivate(ProtViewWindow As ProtectedViewWindow)`  
  Occurs when any Protected View window is activated.
    - `ProtViewWindow As ProtectedViewWindow` (required): The Protected View window that is activated.
- `ProtectedViewWindowDeactivate(ProtViewWindow As ProtectedViewWindow)`  
  Occurs when a Protected View window is deactivated.
    - `ProtViewWindow As ProtectedViewWindow` (required): The deactivated Protected View window.
- `PresentationCloseFinal(Pres As Presentation)`  
  Represents closing the final Presentation object.
    - `Pres As Presentation` (required): The Presentation object.
- `AfterDragDropOnSlide(Sld As Slide, X As Single, Y As Single)`  
  Occurs after an object with the clipboard format "PowerPoint Drop Trigger" has been dropped onto a slide in an open presentation.
    - `Sld As Slide` (required): The slide that raised the event (that is, had a shape added to it).
- `AfterShapeSizeChange(shp As Shape)`  
  Occurs after an object (shape, picture, text box, chart, SmartArt as examples) has been resized on the slide.
    - `shp As Shape` (required): The shape that was resized.
