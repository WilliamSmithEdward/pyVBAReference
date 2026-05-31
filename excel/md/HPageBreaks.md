# HPageBreaks

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024404-0000-0000-C000-000000000046}  

The collection of horizontal page breaks within the print area.

**Remarks:** Each horizontal page break is represented by an HPageBreak object. If you add a page break that does not intersect the print area, the newly-added HPageBreak object will not appear in the HPageBreaks collection for the print area. The contents of the collection may change if the print area is resized or redefined. When the Application property, Count property, Item property, Parent property, or Add method is used in conjunction with the HPageBreaks property of the Worksheet object: - For an automatic print area, the HPageBreaks property applies only to the page breaks within the print area. - For a user-defined print area of the same range, the HPageBreaks property applies to all of the page breaks.

**Example:**

```vba
ActiveSheet.HPageBreaks.Add Before:=ActiveCell
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
- `Item As HPageBreak  (read-only)`  
  Returns a single object from a collection.
- `_Default As HPageBreak  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Before As Object) As HPageBreak`  
  Adds a horizontal page break.
    - `Before As Object` (required): A Range object. The range above which the new page break will be added.
