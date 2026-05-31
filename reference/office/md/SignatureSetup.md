# SignatureSetup

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD6A1-0000-0000-C000-000000000046}  

Represents the information used to set up a signature packet.

**Example:**

```vba
Dim objSigSetup As SignatureSetup
With objSigSetup
.AllowComments = True
.ShowSignDate = True
.SigningInstructions = "Please sign this document."
.SuggestedSignerEmail = "jdow@example.com"
Next
```

## Properties (12)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SignatureSetup object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SignatureSetup object was created. Read-only.
- `ReadOnly As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the SignatureSetup object is read-only. Read-only.
- `Id As String  (read-only)`  
  Gets the ID of the signature provider for a document. Read-only.
- `SignatureProvider As String  (read-only)`  
  Gets a value identifying an installed signature provider add-in. Read-only.
- `SuggestedSigner As String  (read/write)`  
  Gets or sets the name of the principle signer of the document. Read/write.
- `SuggestedSignerLine2 As String  (read/write)`  
  Gets or sets the second line of suggested signer information (for example, title). Read/write.
- `SuggestedSignerEmail As String  (read/write)`  
  Gets or sets the email address of the signer of the document. Read/write.
- `SigningInstructions As String  (read/write)`  
  Gets or sets the instructions for signing the document. Read/write.
- `AllowComments As Boolean  (read/write)`  
  Gets or sets a Boolean value specifying whether the signer can enter comments in the Sign dialog box. Read/write.
- `ShowSignDate As Boolean  (read/write)`  
  Gets or sets a Boolean value indicating whether the date the document was signed should be displayed. Read/write.
- `AdditionalXml As String  (read/write)`  
  Gets or sets any additional XML information added to the signature during setup. Read/write.
