# Dictionary

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209AD-0000-0000-C000-000000000046}  

Represents a dictionary. Dictionary objects that represent custom dictionaries are members of the Dictionaries collection. Other dictionary objects are returned by properties of the Languages collection; these include the ActiveSpellingDictionary, ActiveGrammarDictionary, ActiveThesaurusDictionary, and ActiveHyphenationDictionary properties.

**Remarks:** Use CustomDictionaries (Index), where Index is an index number or the string name for the dictionary, to return a single Dictionary object that represents a custom dictionary. The following example returns the first dictionary in the collection. The following example returns the dictionary named "MyDictionary." Use the ActiveCustomDictionary property to set the custom spelling dictionary in the collection to which new words are added. If you try to set this property to a dictionary that's not a custom dictionary, an error occurs. Use the Add method to add a new dictionary to the collection of active custom dictionaries. If there is no file with the name specified by FileName, Word creates it. The following example adds "MyCustom.dic" to the collection of custom dictionaries. Remarks Use the Name and Path properties to locate any of the dictionaries. The following example displays a message box that contains the full path for each dictionary. Use the LanguageSpecific property to determine whether the specified custom dictionary can have a specific language assigned to it with the LanguageID property.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Dictionary object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Path As String  (read-only)`  
  Returns the path to the specified dictionary. Read-only String.
- `LanguageID As WdLanguageID  (read/write)`  
  Returns or sets a WdLanguageID constant that represents the language for the specified object. Read/write.
- `ReadOnly As Boolean  (read-only)`  
  True if the specified dictionary cannot be changed. Read-only Boolean.
- `Type As WdDictionaryType  (read-only)`  
  Returns the dictionary type. Read-only WdDictionaryType.
- `LanguageSpecific As Boolean  (read/write)`  
  True if the custom dictionary is to be used only with text formatted for a specific language. Read/write Boolean.

## Methods (1)

- `Delete()`  
  Deletes the specified dictionary.
