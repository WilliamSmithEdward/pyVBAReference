# CustomerData

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934F6-5A91-11CF-8700-00AA0060263B}  

Stores information about a customer (such as name, address, telephone number, and so on) or other information in XML form, as a collection of CustomXMLPart objects associated with a Microsoft PowerPoint object.

**Remarks:** You can store customer data in CustomLayout, Master, Presentation, Shape, and Slide objects. You can associate one or more CustomXMLPart objects with the same object. - Customer data persists from one instance to the next in a PowerPoint document only when you save the document in XML file format, as a PowerPoint XML presentation. Customer data does not persist in documents saved in .ppt, .htm, or .mht formats. - There is no user interface associated with customer data in PowerPoint. The only way that you can assign and manipulate customer data is programmatically. Use the Add method to add a new CustomXMLPart object to the CustomerData collection. Use the Delete method to delete a CustomXMLPart object from the CustomerData collection. Use the Item method to get a specific CustomXMLPart object from the collection. Individual items in the collection are represented by GUIDs (globally unique identifiers). Use customer data in the same way that you used Tags objects in versions of PowerPoint previous to Microsoft Office PowerPoint 2007--that is, to associate data with objects.

**Example:**

```vba
Public Sub CustomerData_Example()

    Dim pptCustomXMLPart As CustomXMLPart

    Set pptCustomXMLPart = ActivePresentation.Slides(1).Shapes(1).customerData.Add

    Debug.Print pptCustomXMLPart.Id

    pptCustomXMLPart.LoadXML ("<Customer><CustomerID>Customer #1</CustomerID></Customer>")

    Debug.Print pptCustomXMLPart.xml

End Sub
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of CustomXMLPart objects in the CustomerData collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified CustomerData object. Read-only.
- `Parent As Object  (read-only)`  
  Returns the parent object of the specified CustomerData object. Read-only.

## Methods (3)

- `Item(Id As String) As CustomXMLPart`  
  Returns the specified CustomXMLPart object from the CustomerData collection. Read-only.
    - `Id As String` (required): The ID of the CustomXMLPart object.
- `Add() As CustomXMLPart`  
  Adds a CustomXMLPart to the CustomerData collection of a CustomLayout, Master, Presentation, Shape, or Slide object and returns the CustomXMLPart object created.
- `Delete(Id As String)`  
  Deletes the specified CustomXMLPart object from the CustomerData collection of a CustomLayout, Master, Presentation, Shape, or Slide object.
    - `Id As String` (required): The ID of the CustomXMLPart object to be deleted.
