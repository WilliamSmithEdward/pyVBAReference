# MsoEnvelope

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {0006F01A-0000-0000-C000-000000000046}  

Provides access to functionality that lets you send documents as email messages directly from Microsoft Office applications.

**Remarks:** Use the MailEnvelope property of the Document object, Chart object, or Worksheet object (depending on the application you are using) to return an MsoEnvelope object.

**Example:**

```vba
Sub SendMail(ByVal strRecipient As String)

 'Use a With...End With block to reference the MsoEnvelope object.
 With Application.ActiveDocument.MailEnvelope

 'Add some introductory text before the body of the email.
 .Introduction = "Please read this and send me your comments."

 'Return a Microsoft Outlook MailItem object that
 'you can use to send the document.
 With .Item

 'All of the mail item settings are saved with the document.
 'When you add a recipient to the Recipients collection
 'or change other properties, these settings persist.
 .Recipients.Add strRecipient
 .Subject = "Here is the document."

 'The body of this message will be
 'the content of the active document.
 .Send
 End With
 End With
End Sub
```

## Properties (4)

- `Introduction As String  (read/write)`  
  Sets or gets the introductory text that is included with a document that is sent by using the MsoEnvelope object. The introductory text is included at the top of the document in the email. Read/write.
- `Item As Object  (read-only)`  
  Gets a MailItem object that can be used to send the document as an email. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the MsoEnvelope object. Read-only.
- `CommandBars As Object  (read-only)`  
  Gets a CommandBars collection. Read-only.

## Events (2)

- `EnvelopeShow()`  
  Occurs when the user interface (UI) that corresponds to the MsoEnvelope object is displayed.
- `EnvelopeHide()`  
  Occurs when the user interface (UI) that corresponds to the MsoEnvelope object is hidden.
