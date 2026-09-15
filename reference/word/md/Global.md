# Global

**Type:** Class  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209F0-0000-0000-C000-000000000046}  

Contains top-level properties and methods that don't need to be preceded by the Application property.

**Remarks:** The following two statements have the same result. One statement uses the Application property to access the Documents collection, and one does not. Both statements are equal and achieve the same result.

## Properties (47)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Global object.
- `Name As String  (read-only)`  
  Returns name of the specified object. Read-only String.
- `Documents As Documents  (read-only)`  
  Returns a Documents collection that represents all the open documents. Read-only.
- `Windows As Windows  (read-only)`  
  Returns a Windows collection that represents all open document windows. Read-only.
- `ActiveDocument As Document  (read-only)`  
  Returns a Document object that represents the active document (the document with the focus). Read-only.
- `ActiveWindow As Window  (read-only)`  
  Returns a Window object that represents the active window (the window with the focus). Read-only.
- `Selection As Selection  (read-only)`  
  Returns a Selection object that represents a selected range or the insertion point. Read-only.
- `WordBasic As Object  (read-only)`  
  Returns an Automation object (Word.Basic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.
- `PrintPreview As Boolean  (read/write)`  
  True if print preview is the current view. Read/write Boolean.
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
- `FileConverters As FileConverters  (read-only)`  
  Returns a FileConverters collection that represents all the file converters available to Microsoft Word. Read-only.
- `Dialogs As Dialogs  (read-only)`  
  Returns a Dialogs collection that represents all the built-in dialog boxes in Word. Read-only.
- `CaptionLabels As CaptionLabels  (read-only)`  
  Returns a CaptionLabels collection that represents all the available caption labels. Read-only.
- `AutoCaptions As AutoCaptions  (read-only)`  
  Returns an AutoCaptions collection that represents the captions that are automatically added when items such as tables and pictures are inserted into a document. Read-only.
- `AddIns As AddIns  (read-only)`  
  Returns an AddIns collection that represents all available add-ins, regardless of whether they're currently loaded. Read-only.
- `Tasks As Tasks  (read-only)`  
  Returns a Tasks collection that represents all the applications that are running.
- `MacroContainer As Object  (read-only)`  
  Returns a Template or Document object that represents the template or document in which the module that contains the running procedure is stored.
- `CommandBars As CommandBars  (read-only)`  
  Returns a CommandBars collection that represents the menu bar and all the toolbars in Microsoft Word.
- `SynonymInfo As SynonymInfo  (read-only)`  
  Returns a SynonymInfo object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the specified word or phrase.
- `VBE As VBE  (read-only)`  
  Returns a VBE object that represents the Visual Basic Editor.
- `ListGalleries As ListGalleries  (read-only)`  
  Returns a ListGalleries collection that represents the three list template galleries (Bulleted, Numbered, and Outline Numbered).
- `ActivePrinter As String  (read/write)`  
  Returns or sets the name of the active printer. Read/write String.
- `Templates As Templates  (read-only)`  
  Returns a Templates collection that represents all the available templates-global templates and those attached to open documents.
- `CustomizationContext As Object  (read/write)`  
  Returns or sets a Template or Document object that represents the template or document in which changes to menu bars, toolbars, and key bindings are stored. Read/write. .
- `KeyBindings As KeyBindings  (read-only)`  
  Returns a KeyBindings collection that represents customized key assignments, which include a key code, a key category, and a command.
- `KeysBoundTo As KeysBoundTo  (read-only)`  
  Returns a KeysBoundTo object that represents all the key combinations assigned to the specified item.
- `FindKey As KeyBinding  (read-only)`  
  Returns a KeyBinding object that represents the specified key combination. Read-only.
- `Options As Options  (read-only)`  
  Returns an Options object that represents application settings in Microsoft Word.
- `CustomDictionaries As Dictionaries  (read-only)`  
  Returns a Dictionaries object that represents the collection of active custom dictionaries. Read-only.
- `StatusBar As String  (write-only)`  
  This property is no longer supported in Microsoft Word Visual Basic for Applications.
- `ShowVisualBasicEditor As Boolean  (read/write)`  
  True if the Visual Basic Editor window is visible. Read/write Boolean.
- `IsObjectValid As Boolean  (read-only)`  
  True if the specified variable that references an object is valid. Read-only Boolean.
- `HangulHanjaDictionaries As HangulHanjaConversionDictionaries  (read-only)`  
  Returns a HangulHanjaConversionDictionaries collection that represents all the active custom conversion dictionaries.
- `LanguageSettings As LanguageSettings  (read-only)`  
  Returns a LanguageSettings object, which contains information about the language settings in Microsoft Word.
- `AutoCorrectEmail As AutoCorrect  (read-only)`  
  Returns an AutoCorrect object that represents automatic corrections made to email messages.
- `ProtectedViewWindows As ProtectedViewWindows  (read-only)`  
  Returns a ProtectedViewWindows object that represents the open Protected View windows. Read-only.
- `ActiveProtectedViewWindow As ProtectedViewWindow  (read-only)`  
  Returns a ProtectedViewWindow object that represents the active Protected View window (the Protected View window with the focus). Read-only.
- `IsSandboxed As Boolean  (read-only)`  
  True if the application window is a Protected View window. Read-only.

## Methods (27)

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
  Uses an open dynamic data exchange (DDE) channel to request information from the receiving application, and returns the information as a string.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
    - `Item As String` (required): The item to be requested.
- `DDETerminate(Channel As Long)`  
  Closes the specified dynamic data exchange (DDE) channel to another application.
    - `Channel As Long` (required): The channel number returned by the DDEInitiate method.
- `DDETerminateAll()`  
  Closes all dynamic data exchange (DDE) channels opened by Microsoft Word. .
- `BuildKeyCode(Arg1 As WdKey, [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant]) As Long`  
  Returns a unique number for the specified key combination.
    - `Arg1 As WdKey` (required): A key you specify by using one of the WdKey constants.
    - `Arg2 As Variant` (optional): A key you specify by using one of the WdKey constants.
    - `Arg3 As Variant` (optional): A key you specify by using one of the WdKey constants.
    - `Arg4 As Variant` (optional): A key you specify by using one of the WdKey constants.
- `KeyString(KeyCode As Long, [KeyCode2 As Variant]) As String`  
  Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A).
    - `KeyCode As Long` (required): A key you specify by using one of the WdKey constants.
    - `KeyCode2 As Variant` (optional): A second key you specify by using one of the WdKey constants.
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
- `GetSpellingSuggestions(Word As String, [CustomDictionary As Variant], [IgnoreUppercase As Variant], [MainDictionary As Variant], [SuggestionMode As Variant], [CustomDictionary2 As Variant], [CustomDictionary3 As Variant], [CustomDictionary4 As Variant], [CustomDictionary5 As Variant], [CustomDictionary6 As Variant], [CustomDictionary7 As Variant], [CustomDictionary8 As Variant], [CustomDictionary9 As Variant], [CustomDictionary10 As Variant]) As SpellingSuggestions`  
  Returns a SpellingSuggestions collection that represents the words suggested as spelling replacements for a given word.
    - `Word As String` (required): The word whose spelling is to be checked.
    - `IgnoreUppercase As Variant` (optional): True to ignore words in all uppercase letters. If this argument is omitted, the current value of the IgnoreUppercase property is used.
    - `SuggestionMode As Variant` (optional): Specifies the way Word makes spelling suggestions. Can be one of the following WdSpellingWordType constants: wdAnagram, wdSpellword, or wdWildcard. The default value is WdSpellword.
- `Help(HelpType As Variant)`  
  Displays on-line Help information.
    - `HelpType As Variant` (required): The on-line Help topic or window. Can be any of these WdHelpType constants.
- `NewWindow() As Window`  
  Opens a new window with the same document as the specified window. Returns a Window object.
- `CleanString(String As String) As String`  
  Removes nonprinting characters (character codes 1&ndash;29) and special Word characters from the specified string or changes them to spaces (character code 32), as described in the "Remarks" section. Returns the result as a String.
    - `String As String` (required): The source string.
- `ChangeFileOpenDirectory(Path As String)`  
  Sets the folder in which Word searches for documents. .
    - `Path As String` (required): The path to the folder in which Word searches for documents.
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
- `PointsToPixels(Points As Single, [fVertical As Variant]) As Single`  
  Converts a measurement from points to pixels. Returns the converted measurement as a Single.
    - `Points As Single` (required): The point value to be converted to pixels.
    - `fVertical As Variant` (optional): True to return the result as vertical pixels; False to return the result as horizontal pixels.
- `PixelsToPoints(Pixels As Single, [fVertical As Variant]) As Single`  
  Converts a measurement from pixels to points. Returns the converted measurement as a Single.
    - `Pixels As Single` (required): The pixel value to be converted to points.
    - `fVertical As Variant` (optional): True to convert vertical pixels; False to convert horizontal pixels.
