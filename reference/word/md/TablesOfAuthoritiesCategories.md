# TablesOfAuthoritiesCategories

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020976-0000-0000-C000-000000000046}  

A collection of TableOfAuthoritiesCategory objects that represent the table of authorities categories, such as Cases and Statutes. The TablesOfAuthoritiesCategories collection includes all 16 categories listed in the Category box on the Table of Authorities tab in the Index and Tables dialog box.

**Remarks:** Use the TablesOfAuthoritiesCategories property to return the TablesOfAuthoritiesCategories collection. The following example displays the names of the categories in the TablesOfAuthoritiesCategories collection. The Add method isn't available for the TablesOfAuthoritiesCategories collection. The collection is limited to 16 items; however, you can use the Name property to rename an existing category. Use TablesOfAuthoritiesCategories (Index), where Index is the category name or index number, to return a single TableOfAuthoritiesCategory object. The following example renames the Rules category as Other Provisions. The index number represents the position of the category in the Index and Tables dialog box. The following example displays the name of the first category in the TablesOfAuthoritiesCategories collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TablesOfAuthoritiesCategories object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of categories in the collection. Read-only.

## Methods (1)

- `Item(Index As Variant) As TableOfAuthoritiesCategory`  
  Returns an individual TablesOfAuthoritiesCategory object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
