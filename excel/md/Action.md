# Action

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002448F-0000-0000-C000-000000000046}  

Represents an action to be executed in a PivotTable or sheet data.

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Caption As String  (read-only)`  
  Returns the caption assigned to the Action object. Read-only String.
- `Type As XlActionType  (read-only)`  
  Returns the action type. Read-only XlActionType.
- `Coordinate As String  (read-only)`  
  Returns the coordinate property of the Action object. Read-only.
- `Content As String  (read-only)`  
  Returns the content associated to the Action object. Read-only String.

## Methods (1)

- `Execute()`  
  Performs the specified action.
