# TableOfAuthoritiesCategory

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020977-0000-0000-C000-000000000046}  

Represents a single table of authorities category. The TableOfAuthoritiesCategories object is a member of the TablesOfAuthoritiesCategories collection.

**Remarks:** The TablesOfAuthoritiesCategories collection includes all 16 categories listed in the Category box on the Table of Authorities tab in the Index and Tables dialog box. Use TablesOfAuthoritiesCategories (Index), where Index is the category name or index number, to return a single TableOfAuthoritiesCategory object. The following example renames the Rules category as Other Provisions. The index number represents the position of the category in the Index and Tables dialog box (Insert menu). The following example displays the name of the first category in the TablesOfAuthoritiesCategories collection. The Add method isn't available for the TablesOfAuthoritiesCategories collection. The collection is limited to 16 items; however, you can use the Name property to rename an existing category.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TableOfAuthoritiesCategory object.
- `Name As String  (read/write)`  
  Returns the name of the specified object. Read-only String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
