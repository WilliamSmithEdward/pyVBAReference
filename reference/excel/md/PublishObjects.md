# PublishObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024443-0000-0000-C000-000000000046}  

A collection of all PublishObject objects in the workbook.

**Remarks:** Each PublishObject object represents an item in a workbook that has been saved to a webpage and can be refreshed according to values specified by the properties and methods of the object.

**Example:**

```vba
Set objPObjs = ActiveWorkbook.PublishObjects
For Each objPO in objPObjs
 If objPO.HtmlType = xlHTMLStatic Then
 objPO.Publish
 End If
Next objPO
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
- `Item As PublishObject  (read-only)`  
  Returns a single object from a collection.
- `_Default As PublishObject  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Add(SourceType As XlSourceType, Filename As String, [Sheet As Variant], [Source As Variant], [HtmlType As Variant], [DivID As Variant], [Title As Variant]) As PublishObject`  
  Creates an object that represents an item in a document saved to a webpage. Such objects facilitate subsequent updates to the webpage while automated changes are being made to the document in Microsoft Excel. Returns a PublishObject object.
    - `SourceType As XlSourceType` (required): The source type.
    - `Filename As String` (required): String. The URL (on the intranet or the web) or path (local or network) to which the source object was saved.
    - `Sheet As Variant` (optional): The name of the worksheet that was saved as a webpage.
    - `Source As Variant` (optional): A unique name used to identify items that have one of the following constants as their _SourceType_ argument: xlSourceAutoFilter, xlSourceChart, xlSourcePivotTable, xlSourcePrintArea, xlSourceQuery, or xlSourceRange. If _SourceType_ is xlSourceRange, _Source_ specifies a range, which can be a defined name. If _SourceType_ is xlSourceChart, xlSourcePivotTable, or xlSourceQuery, _Source_ specifies the name of a chart, PivotTable report, or query table.
    - `HtmlType As Variant` (optional): Specifies whether the item is saved as an interactive Microsoft Office Web component or as static text and images. Can be one of the XlHTMLType constants: xlHtmlCalc, xlHtmlChart, xlHtmlList, or xlHtmlStatic.
    - `DivID As Variant` (optional): The unique identifier used in the HTML DIV tag to identify the item on the webpage.
    - `Title As Variant` (optional): The title of the webpage.
- `Delete()`  
  Deletes the object.
- `Publish()`  
  Saves a copy of the item or items in the spreadsheet that have been added to the PublishObjects collection to a webpage.
