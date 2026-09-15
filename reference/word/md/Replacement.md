# Replacement

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B1-0000-0000-C000-000000000046}  

Represents the replace criteria for a find-and-replace operation. The properties and methods of the Replacement object correspond to the options in the Find and Replace dialog box.

**Remarks:** Use the Replacement property to return a Replacement object. The following example replaces the next occurrence of the word "hi" with the word "hello." To find and replace formatting, set both the find text and the replace text to empty strings ("") and set the Format argument of the Execute method to True. The following example removes all the bold formatting in the active document. The Bold property is True for the Find object and False for the Replacement object.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Replacement object.
- `Font As Font  (read/write)`  
  Returns or sets a Font object that represents the character formatting of the specified object. Read/write Font.
- `ParagraphFormat As ParagraphFormat  (read/write)`  
  Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified replacement operation. Read/write.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified object. To set this property, specify the local name of the style, an integer, a WdBuiltinStyle constant, or an object that represents the style. Read/write Variant.
- `Text As String  (read/write)`  
  Returns or sets the text to replace. Read/write String.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets a WdLanguageID constant that represents the language for the specified range. Read/write.
- `Highlight As Long  (read/write)`  
  True if highlight formatting is applied to the replacement text. Read/write Long.
- `Frame As Frame  (read-only)`  
  Returns a Frame object that represents the frame formatting for the specified style or find-and-replace operation. Read-only.
- `LanguageIDFarEast As WdLanguageID  (read/write)`  
  Returns or sets an East Asian language for the specified replacement. Read/write WdLanguageID.
- `NoProofing As Long  (read/write)`  
  True if Microsoft Word replaces text that the spelling and grammar checker ignores. Read/write Long.

## Methods (1)

- `ClearFormatting()`  
  Removes text and paragraph formatting from the text specified in a replace operation.
