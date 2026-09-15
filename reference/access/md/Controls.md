# Controls

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {5970C574-EB8C-11CD-8701-00AA003F0F07}  

The Controls collection contains all of the controls on a form, report, or subform, within another control, or attached to another control. The Controls collection is a member of the Form, Report, and SubForm objects.

**Remarks:** You can enumerate individual controls, count them, and set their properties in the Controls collection. For example, you can enumerate the Controls collection of a particular form and set the Height property of each control to a specified value. It's faster to refer to the Controls collection implicitly, as in the following examples, which refer to a control called NewData on a form named OrderForm. Of the following syntax examples, Me!NewData is the fastest way to refer to the control. You can also refer to an individual control by referring explicitly to the Controls collection. Additionally, you can refer to a control by its index in the collection. The Controls collection is indexed beginning with zero. To work with the controls on a section of a form or report, use the Section property to return a reference to a Section object. You can then refer to the Controls collection of the Section object. Two types of Control objects, the tab control and option group control, have Controls collections that can contain multiple controls.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Object  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
