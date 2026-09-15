# Pages

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E978-E47C-11CD-8701-00AA003F0F07}  

The Pages collection contains all Page objects in a tab control.

**Remarks:** The Pages collection is a special kind of Controls collection belonging to the tab control. It contains Page objects, which are controls. The Pages collection differs from a typical Controls collection in that you can add and remove Page objects by using methods of the Pages collection. To add a new Page object to the Pages collection from Visual Basic, use the Add method. To remove an existing Page object, use the Remove method. To count the number of Page objects in the Pages collection, use the Count property. You can also use the CreateControl method to add a Page object to the Pages collection of a tab control. To do this, you must specify the name of the tab control for the Parent argument of the CreateControl function. The ControlType property constant for a Page object is acPage. You can enumerate through the Pages collection by using the For Each...Next statement. Individual Page objects in the Pages collection are indexed beginning with zero.

## Properties (2)

- `Item As Page  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Page.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.

## Methods (2)

- `Add([Before As Variant]) As Page`  
  The Add method adds a new Page object to the Pages collection of a tab control.
    - `Before As Variant` (optional): An Integer that specifies the index of the Page object before which the new Page object should be added. The index of the Page object corresponds to the value of the PageIndex property for that Page object. If you omit this argument, the new Page object is added to the end of the collection.
- `Remove([Item As Variant])`  
  The Remove method removes a Page object from the Pages collection of a tab control.
    - `Item As Variant` (optional): An integer that specifies the index of the Page object to be removed. The index of the Page object corresponds to the value of the PageIndex property for that Page object. If you omit this argument, the last Page object in the collection is removed.
