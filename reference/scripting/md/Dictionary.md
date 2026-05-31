# Dictionary

**Type:** Class  
**Library:** Microsoft Scripting Runtime  
**GUID:** {EE09B103-97E0-11CF-978F-00A02463E06F}  

Scripting.Dictionary

## Properties (4)

- `Item As Variant  (read/write)`  
  Set or get the item for a given key
- `Count As Long  (read-only)`  
  Get the number of items in the dictionary.
- `Key As Variant  (write-only)`  
  Change a key to a different key.
- `CompareMode As CompareMethod  (read/write)`  
  Set or get the string comparison method.

## Methods (6)

- `Add(Key As Variant, Item As Variant)`  
  Add a new key and item to the dictionary.
- `Exists(Key As Variant) As Boolean`  
  Determine if a given key is in the dictionary.
- `Items() As Variant`  
  Get an array containing all items in the dictionary.
- `Keys() As Variant`  
  Get an array containing all keys in the dictionary.
- `Remove(Key As Variant)`  
  Remove a given key from the dictionary.
- `RemoveAll()`  
  Remove all information from the dictionary.
