# Style

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002092C-0000-0000-C000-000000000046}  

Represents a single built-in or user-defined style. The Style object includes style attributes (such as font, font style, and paragraph spacing) as properties of the Style object. The Style object is a member of the Styles collection. The Styles collection includes all the styles in the specified document.

**Remarks:** Use Styles (Index), where Index is the style name, a WdBuiltinStyle constant or index number, to return a single Style object. You must exactly match the spelling and spacing of the style name, but not necessarily its capitalization. The following example modifies the font name of the user-defined style named "Color" in the active document. The following example sets the built-in Heading 1 style to not be bold. The style index number represents the position of the style in the alphabetically sorted list of style names. Note that Styles(1) is the first style in the alphabetical list. The following example displays the base style and style name of the first style in the Styles collection. To apply a style to a range, paragraph, or multiple paragraphs, set the Style property to a user-defined or built-in style name. The following example applies the Normal style to the first four paragraphs in the active document. The following example applies the Heading 1 style to the first paragraph in the selection. The following example creates a character style named "Bolded" and applies it to the selection. Use the OrganizerCopy method to copy styles between documents and templates.

## Properties (30)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Style object.
- `NameLocal As String  (read/write)`  
  Returns the name of a built-in style in the language of the user. Read/write String.
- `BaseStyle As Variant  (read/write)`  
  Returns or sets an existing style on which you can base the formatting of another style. Read/write Variant.
- `Description As String  (read-only)`  
  Returns the description of the specified style. Read-only String.
- `Type As WdStyleType  (read-only)`  
  Returns the style type. Read-only WdStyleType.
- `BuiltIn As Boolean  (read-only)`  
  True if the specified object is one of the built-in styles or caption labels in Word. Read-only Boolean.
- `NextParagraphStyle As Variant  (read/write)`  
  Returns or sets the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style. Read/write Variant.
- `InUse As Boolean  (read-only)`  
  True if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. Read-only Boolean.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified style.
- `ParagraphFormat As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified style. Read/write.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified style. Read/write Font.
- `Frame As Frame  (read-only)`  
  Returns a Frame object that represents the frame formatting for the specified style. Read-only.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets a WdLanguageID constant that represents the language for the specified range. Read/write.
- `AutomaticallyUpdate As Boolean  (read/write)`  
  True if the style is automatically redefined based on the selection. Read/write Boolean.
- `ListTemplate As ListTemplate  (read-only)`  
  Returns a ListTemplate object that represents the list formatting for the specified Style object.
- `ListLevelNumber As Long  (read-only)`  
  Returns the list level for the specified style. Read-only Long.
- `LanguageIDFarEast As WdLanguageID  (read/write)`  
  Returns or sets an East Asian language for the specified object. Read/write WdLanguageID.
- `NoProofing As Long  (read/write)`  
  True if the spelling and grammar checker ignores text formatted with this style. Read/write Long.
- `LinkStyle As Variant  (read/write)`  
  Sets or returns a Variant that represents a link between a paragraph and a character style. Read/write.
- `Visibility As Boolean  (read/write)`  
  True if the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane. Read/write.
- `NoSpaceBetweenParagraphsOfSameStyle As Boolean  (read/write)`  
  True for Microsoft Word to remove spacing between paragraphs that are formatted using the same style. Read/write Boolean.
- `Table As TableStyle  (read-only)`  
  Returns a TableStyle object representing properties that can be applied to a table using a table style.
- `Locked As Boolean  (read/write)`  
  True if a style cannot be changed or edited. Read/write Boolean.
- `Priority As Long  (read/write)`  
  Returns or sets a Long that represents the priority for sorting styles in the Styles task pane. Read/write.
- `UnhideWhenUsed As Boolean  (read/write)`  
  True if the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Word after it is used in the document. Read/write.
- `QuickStyle As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether the style corresponds to an available quick style. Read/write.
- `Linked As Boolean  (read-only)`  
  Returns a Boolean that represents whether a style is a linked style that can be used for both paragraph and character formatting. Read-only.

## Methods (2)

- `Delete()`  
  Deletes the specified style.
- `LinkToListTemplate(ListTemplate As ListTemplate, [ListLevelNumber As Variant])`  
  Links the specified style to a list template so that the style's formatting can be applied to lists.
    - `ListTemplate As ListTemplate` (required): The list template that the style is to be linked to.
    - `ListLevelNumber As Variant` (optional): An integer corresponding to the list level that the style is to be linked to. If this argument is omitted, then the level of the style is used.
