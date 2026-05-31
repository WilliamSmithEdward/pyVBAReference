# VPageBreaks

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024405-0000-0000-C000-000000000046}  

A collection of vertical page breaks within the print area.

**Remarks:** Each vertical page break is represented by a VPageBreak object. When the Application property, Count property, Creator property, Item property, Parent property, or Add method is used in conjunction with the VPageBreaks property: - For an automatic print area, the VPageBreaks property applies only to the page breaks within the print area. - For a user-defined print area of the same range, the VPageBreaks property applies to all of the page breaks.

**Example:**

```vba
ActiveSheet.VPageBreaks.Add Before:=ActiveCell
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As VPageBreak  (read-only)`  
  Returns a single object from a collection.
- `_Default As VPageBreak  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Before As Object) As VPageBreak`  
  Adds a vertical page break.
    - `Before As Object` (required): A Range object. The range to the left of which the new page break will be added.
