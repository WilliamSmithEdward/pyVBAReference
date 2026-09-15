# TaskPanes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E6AAEC05-E543-4085-BA92-9BF7D2474F5C}  

A collection of TaskPane objects that contains commonly performed tasks in Microsoft Word.

**Remarks:** Use the TaskPanes property to return the TaskPanes collection. Use the Item method with a WdTaskPanes constant to refer to a specific task pane. The example below displays the formatting task pane.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of task panes in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TaskPanes object.

## Methods (1)

- `Item(Index As WdTaskPanes) As TaskPane`  
  Returns the specified task pane as a TaskPane object.
    - `Index As WdTaskPanes` (required): The specified task pane.
