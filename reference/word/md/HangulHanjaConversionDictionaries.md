# HangulHanjaConversionDictionaries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E0-0000-0000-C000-000000000046}  

A collection of Dictionary objects that includes the active custom Hangul-Hanja conversion dictionaries.

**Remarks:** Use the HangulHanjaDictionaries property to return the collection of currently active custom conversion dictionaries. The following example displays the names of all the active custom conversion dictionaries. Use the Add method to add a new custom conversion dictionary to the collection of active custom conversion dictionaries. If there isn't a file with the name specified by FileName, Microsoft Word creates it. The following example adds "Hanja1.hhd" to the collection of custom conversion dictionaries. Use the ClearAll method to unload all custom conversion dictionaries. Note, however, that this method doesn't delete the dictionary files. After you use this method, the number of custom conversion dictionaries in the collection is 0 (zero). The following example clears the custom conversion dictionaries and creates a new custom conversion dictionary file. The new dictionary is set as the active custom dictionary to which Word will automatically add any new words it encounters. You set the custom dictionary to which new words are added by using the ActiveCustomDictionary property.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HangulHanjaConversionDictionaries object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of dictionaries in the collection. Read-only.
- `Maximum As Long  (read-only)`  
  Returns the maximum number of custom or conversion dictionaries allowed. Read-only Long.
- `ActiveCustomDictionary As Dictionary  (read/write)`  
  Returns or sets a Dictionary object that represents the custom dictionary to which words will be added. Read/write.

## Methods (3)

- `Item(Index As Variant) As Dictionary`  
  Returns an individual Dictionary object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(FileName As String) As Dictionary`  
  Returns a Dictionary object that represents a new custom spelling or conversion dictionary added to the collection of active custom spelling or conversion dictionaries.
    - `FileName As String` (required): The string name of the dictionary file. If no path is specified in the string, the proofing tools path is used.
- `ClearAll()`  
  Unloads all of the custom or conversion dictionaries.
