# Pages

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A3-0000-0000-C000-000000000046}  

A collection of pages in a document. Use the Pages collection and the related objects and properties for programmatically defining page layout in a workbook.

**Remarks:** Use the Pages property of the PageSetup object to return a Pages collection. The following example accesses all pages on the active worksheet. Use the Item method to access an individual Page object that represents an individual page on a worksheet. The following example accesses the first page on the active worksheet.

## Properties (4)

- `Item As Page  (read-only)`  
  Returns a Page object that represents a collection of pages in a workbook. Read-only.
- `_Default As Page  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `_NewEnum As IUnknown  (read-only)`
