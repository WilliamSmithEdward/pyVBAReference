# MailingLabel

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020917-0000-0000-C000-000000000046}  

Represents a mailing label.

**Remarks:** Use the MailingLabel property to return the MailingLabel object. The following example sets default mailing label options. Use the PrintOut method to print a mailing label listed in the Product Number box in the Label Options dialog box. The following example prints a page of Avery 5162 standard address labels using the specified address. Use the CustomLabels property to format or print a custom mailing label. The following example sets the number of labels across and down for the custom label named "MyLabel."

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailingLabel object.
- `DefaultLaserTray As WdPaperTray  (read/write)`  
  Returns or sets the default paper tray that contains sheets of mailing labels. Read/write WdPaperTray.
- `CustomLabels As CustomLabels  (read-only)`  
  Returns a CustomLabels collection that represents the available custom mailing labels. Read-only.
- `DefaultLabelName As String  (read/write)`  
  Returns or sets the name for the default mailing label. Read/write String.
- `Vertical As Boolean  (read/write)`  
  True vertically orients text on Asian mailing labels. Read/write Boolean.

## Methods (5)

- `LabelOptions()`  
  Displays the Label Options dialog box.
- `CreateNewDocument([Name As Variant], [Address As Variant], [AutoText As Variant], [ExtractAddress As Variant], [LaserTray As Variant], [PrintEPostageLabel As Variant], [Vertical As Variant]) As Document`  
  Creates a new label document using either the default label options or ones that you specify. Returns a Document object that represents the new document.
    - `Name As Variant` (optional): The mailing label name.
    - `Address As Variant` (optional): The text for the mailing label.
    - `AutoText As Variant` (optional): The name of the AutoText entry that includes the mailing label text.
    - `ExtractAddress As Variant` (optional): True to use the address text marked by the user-defined bookmark named "EnvelopeAddress" instead of using the Address argument.
    - `LaserTray As Variant` (optional): The laser printer tray. Can be one of the WdPaperTray constants.
    - `PrintEPostageLabel As Variant` (optional): True to print postage using an Internet e-postage vendor.
    - `Vertical As Variant` (optional): True formats text vertically on the label. Used for Asian-language mailing labels.
- `PrintOut([Name As Variant], [Address As Variant], [ExtractAddress As Variant], [LaserTray As Variant], [SingleLabel As Variant], [Row As Variant], [Column As Variant], [PrintEPostageLabel As Variant], [Vertical As Variant])`  
  Prints a label or a page of labels with the same address.
    - `Name As Variant` (optional): The mailing label name.
    - `Address As Variant` (optional): The text for the label address.
    - `ExtractAddress As Variant` (optional): True to use the text marked by the "EnvelopeAddress" bookmark (a user-defined bookmark) as the label text. If this argument is specified, Address and AutoText are ignored.
    - `LaserTray As Variant` (optional): The laser printer tray to be used. Can be any WdPaperTray constant.
    - `SingleLabel As Variant` (optional): True to print a single label; False to print an entire page of the same label.
    - `Row As Variant` (optional): The label row for a single label. Not valid if SingleLabel is False.
    - `Column As Variant` (optional): The label column for a single label. Not valid if SingleLabel is False.
    - `PrintEPostageLabel As Variant` (optional): True to print postage using an Internet e-postage vendor.
    - `Vertical As Variant` (optional): True prints text vertically on the label. Used for Asian-language mailing labels.
- `CreateNewDocumentByID([LabelID As Variant], [Address As Variant], [AutoText As Variant], [ExtractAddress As Variant], [LaserTray As Variant], [PrintEPostageLabel As Variant], [Vertical As Variant]) As Document`  
  Creates a new label document using either the default label options or ones that you specify. Returns a Document object that represents the new document.
    - `LabelID As Variant` (optional): The mailing label identification.
    - `Address As Variant` (optional): The text for the mailing label.
    - `AutoText As Variant` (optional): The name of the AutoText entry that includes the mailing label text.
    - `ExtractAddress As Variant` (optional): True to use the address text marked by the user-defined bookmark named "EnvelopeAddress" instead of using the Address argument.
    - `LaserTray As Variant` (optional): The laser printer tray. Can be one of the WdPaperTray constants.
    - `PrintEPostageLabel As Variant` (optional): True to print postage using an Internet e-postage vendor.
    - `Vertical As Variant` (optional): True formats text vertically on the label. Used for Asian-language mailing labels.
- `PrintOutByID([LabelID As Variant], [Address As Variant], [ExtractAddress As Variant], [LaserTray As Variant], [SingleLabel As Variant], [Row As Variant], [Column As Variant], [PrintEPostageLabel As Variant], [Vertical As Variant])`  
  Prints a label or a page of labels with the same address.
    - `LabelID As Variant` (optional): The mailing label identification.
    - `Address As Variant` (optional): The text for the label address.
    - `ExtractAddress As Variant` (optional): True to use the text marked by the "EnvelopeAddress" bookmark (a user-defined bookmark) as the label text. If this argument is specified, Address and AutoText are ignored.
    - `LaserTray As Variant` (optional): The laser printer tray to be used. Can be any WdPaperTray constant.
    - `SingleLabel As Variant` (optional): True to print a single label; False to print an entire page of the same label.
    - `Row As Variant` (optional): The label row for a single label. Not valid if SingleLabel is False.
    - `Column As Variant` (optional): The label column for a single label. Not valid if SingleLabel is False.
    - `PrintEPostageLabel As Variant` (optional): True to print postage using an Internet e-postage vendor.
    - `Vertical As Variant` (optional): True prints text vertically on the label. Used for Asian-language mailing labels.
