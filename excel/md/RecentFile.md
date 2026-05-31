# RecentFile

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024407-0000-0000-C000-000000000046}  

Represents a file in the list of recently used files.

**Remarks:** The RecentFile object is a member of the RecentFiles collection.

**Example:**

```vba
Application.RecentFiles(2).Open
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Path As String  (read-only)`  
  Returns a String value that represents the complete path to the workbook or file that this RecentFile object represents.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.

## Methods (2)

- `Open() As Workbook`  
  Opens a recent workbook.
- `Delete()`  
  Deletes the object.
