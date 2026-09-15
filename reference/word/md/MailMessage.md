# MailMessage

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209BA-0000-0000-C000-000000000046}  

Represents the active email message if you are using Microsoft Word as your email editor.

**Remarks:** Use the MailMessage property to return the MailMessage object. The following example validates the email addresses that appear in the active email message. The methods of the MailMessage object require that you are using Word as your email editor and that an email message is active. If either of these conditions is not true, an error occurs.

## Properties (3)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified MailMessage object.

## Methods (11)

- `CheckName()`  
  Validates the email addresses that appear in the To, Cc, and Bcc lines in the active email message.
- `Delete()`  
  Deletes the specified mail message.
- `DisplayMoveDialog()`  
  Displays the Move dialog box, in which the user can specify a new location for the active email message in an available message store.
- `DisplayProperties()`  
  Displays the Properties dialog box for the active email message.
- `DisplaySelectNamesDialog()`  
  Displays the Select Names dialog box, in which the user can add addresses to the To, Cc, and Bcc lines in the active, unsent email message.
- `Forward()`  
  Opens a new email message with an empty To line for forwarding the active message.
- `GoToNext()`  
  Displays the next mail message if you are using Word as your email editor.
- `GoToPrevious()`  
  Displays the previous mail message if you are using Word as your email editor.
- `Reply()`  
  Opens a new email message - with the sender's address on the To line - for replying to the active message.
- `ReplyAll()`  
  Opens a new email message - with the sender's and all other recipients' addresses on the To and Cc lines, as appropriate - for replying to the active message.
- `ToggleHeader()`  
  Toggles the display of the header in the active email message.
