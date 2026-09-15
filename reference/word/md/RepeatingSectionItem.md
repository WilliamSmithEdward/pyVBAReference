# RepeatingSectionItem

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {4265ED97-A922-4CA4-8CD8-99684CCA9CDB}  

Represents a repeating section item in a content control.

**Remarks:** To get a RepeatingSectionItem object, use the Item method of the RepeatingSectionItemColl collection. You can insert additional repeating section items before or after the specified repeating section item by using the InsertItemBefore or InsertItemAfter methods. To delete the specified repeating section item, use the Delete method. To get the range of the specified repeating section item, use the Range property.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified RepeatingSectionItem object.
- `Range As Range  (read-only)`  
  Returns the range of the specified repeating section item, excluding the start and end tags. Read-only.

## Methods (3)

- `InsertItemBefore() As RepeatingSectionItem`  
  Adds a repeating section item before the specified item and returns the new item.
- `InsertItemAfter() As RepeatingSectionItem`  
  Adds a repeating section item after the specified item and returns the new item.
- `Delete()`  
  Deletes the specified repeating section item.
