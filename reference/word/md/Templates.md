# Templates

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A2-0000-0000-C000-000000000046}  

A collection of Template objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the Templates and Add-ins dialog box.

**Remarks:** Use the Templates property to return the Templates collection. The following example displays the path and file name of each template in the Templates collection. The Add method isn't available for the Templates collection. Instead, you can add a template to the Templates collection by doing any of the following: - Using the Open method with the Documents collection to open a document based on a template or a template - Using the Add method with the Documents collection to open a new document based on a template - Using the Add method with the Addins collection to load a global template - Using the AttachedTemplate property with the Document object to attach a template to a document Use Templates (Index), where Index is the template name or the index number, to return a single Template object. The following example saves the Dot1.dot template. The index number represents the position of the template in the Templates collection. The following example displays the file name of the first template in the Templates collection. Use the NormalTemplate property to return a template object that refers to the Normal template.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Templates object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of templates in the specified collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As Template`  
  Returns an individual Template object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `LoadBuildingBlocks()`  
  Loads the building blocks for all templates into Microsoft Word.
