# SignatureSet

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0410-0000-0000-C000-000000000046}  

A collection of Signature objects that correspond to the digital signature attached to a document.

**Remarks:** Use the Signatures property of the Document object to return a SignatureSet collection; for example: You can add a Signature object to a SignatureSet collection by using the Add method, and you can return an existing member by using the Item method. The AddSignatureLine method also adds a Signature object to the collection. Also see the Subset property, which acts as a filter for whether certain Signature objects appear in the collection. To remove a Signature from a SignatureSet collection, use the Delete method of the Signature object.

**Example:**

```vba
Function AddSignature(ByVal strIssuer As String, _
 strSigner As String) As Boolean

 Dim sig As Signature

 'Display the dialog box that lets the
 'user select a digital signature.
 'If the user selects a signature, then
 'it is added to the Signatures
 'collection. If the user doesn't, then
 'an error is returned.
 Set sig = ActiveDocument.Signatures.Add

 'Test several properties before committing the Signature object to disk.
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

End Function
```

## Properties (9)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SignatureSet object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SignatureSet object was created. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SignatureSet object. Read-only.
- `Item As Signature  (read-only)`  
  Gets a Signature object that corresponds to one of the digital signatures with which the document is currently signed. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SignatureSet object. Read-only.
- `CanAddSignatureLine As Boolean  (read-only)`  
  Gets a Boolean value indicating whether you can add a signature line to a document. Read-only.
- `Subset As MsoSignatureSubset  (read/write)`  
  Gets or sets a value that acts as a filter on the available Signature objects for a document. Read/write.
- `ShowSignaturesPane As Boolean  (write-only)`  
  Gets or sets a Boolean value indicating whether the Signature task pane should be displayed. Read/write.

## Methods (2)

- `AddNonVisibleSignature([varSigProv As Variant]) As Signature`  
  Creates a signature packet when digitally signing a document.
    - `varSigProv As Variant` (optional): Represents the ID of the signature provider.
- `AddSignatureLine([varSigProv As Variant]) As Signature`  
  Adds lines to a document where signatures are collected.
    - `varSigProv As Variant` (optional): Represents the ID of the signature provider.
