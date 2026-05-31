# Signature

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0411-0000-0000-C000-000000000046}  

Represents a digital signature attached to a document. Signature objects are contained in the SignatureSet collection of the Document object.

**Remarks:** You can add a Signature object to a SignatureSet collection by using the Add method, and you can return an existing member by using the Item method. To remove a Signature from a SignatureSet collection, use the Delete method of the Signature object.

**Example:**

```vba
Function AddSignature(ByVal strIssuer As String, _
 strSigner As String) As Boolean

 On Error GoTo Error_Handler

 Dim sig As Signature

 'Display the dialog box that lets the
 'user select a digital signature.
 'If the user selects a signature, then
 'it is added to the Signatures
 'collection. If the user does not, then
 'an error is returned.
 Set sig = ActiveDocument.Signatures.Add

 'Test several properties before commiting the Signature object to disk.
 If sig.Issuer = strIssuer And _
 sig.Signer = strSigner And _
 sig.IsCertificateExpired = False And _
 sig.IsCertificateRevoked = False And _
 sig.IsValid = True Then

 MsgBox "Signed"
 AddSignature = True
 'Otherwise, remove the Signature object from the SignatureSet collection.
 Else
 sig.Delete
 MsgBox "Not signed"
 AddSignature = False
 End If

 'Commit all signatures in the SignatureSet collection to the disk.
 ActiveDocument.Signatures.Commit

 Exit Function
Error_Handler:
 AddSignature = False
 MsgBox "Action canceled."
End Function
```

## Properties (10)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the Signature object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the Signature object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the Signature object. Read-only.
- `IsSigned As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the document was signed successfully. Read-only.
- `Details As SignatureInfo  (read-only)`  
  Gets information about a signature. Read-only.
- `CanSetup As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the user can set properties of the Signature object. Read-only.
- `Setup As SignatureSetup  (read-only)`  
  Gets a SignatureSetup object that provides access to various properties of a signature packet. Read-only.
- `IsSignatureLine As Boolean  (read-only)`  
  Gets a value indicating whether this is a signature line. Read-only.
- `SignatureLineShape As Object  (read-only)`  
  Gets the Shape object associated with a Signature object that is a signature line. Read-only.
- `SortHint As Long  (read-only)`  
  Gets a value representing the sort order of the signatures in a packet with multiple signatures. Read-only.

## Methods (3)

- `Delete()`  
  Deletes the Signature object from the collection.
- `Sign([varSigImg As Variant], [varDelSuggSigner As Variant], [varDelSuggSignerLine2 As Variant], [varDelSuggSignerEmail As Variant])`  
  Creates a signature packet.
    - `varSigImg As Variant` (optional): The signature line graphic image.
    - `varDelSuggSigner As Variant` (optional): The suggested signer.
    - `varDelSuggSignerLine2 As Variant` (optional): The additional signature line.
    - `varDelSuggSignerEmail As Variant` (optional): The email address of the suggested signer.
- `ShowDetails()`  
  Displays details related to a signature packet.
