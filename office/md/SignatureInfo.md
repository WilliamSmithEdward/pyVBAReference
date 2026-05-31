# SignatureInfo

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD6A2-0000-0000-C000-000000000046}  

Represents the information used to create a digital or in-document signature.

**Example:**

```vba
Sub GetCertDetails()
Dim objSignatureInfo As SignatureInfo
Dim varDetail As Variant

strDetail = objSignatureInfo.GetCertificateDetail(certdetExpirationDate)

End Sub
```

## Properties (13)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SignatureInfo object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SignatureInfo object was created. Read-only.
- `ReadOnly As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the SignatureInfo object is read-only. Read-only.
- `SignatureProvider As String  (read-only)`  
  Gets a value identifying an installed signature provider add-in. Read-only.
- `SignatureText As String  (read/write)`  
  Gets or sets the value of the signature text used to sign this document. Read/write.
- `SignatureImage As IPictureDisp  (read/write)`  
  Gets or sets the value of the image used to sign the document. Read/write.
- `SignatureComment As String  (read/write)`  
  Gets or sets a value containing comments included in a signature packet. Read/write.
- `ContentVerificationResults As ContentVerificationResults  (read-only)`  
  Gets a value representing the results of the verification of the hashed contents of a signed document. Read-only.
- `CertificateVerificationResults As CertificateVerificationResults  (read-only)`  
  Gets the results from the verification of a digital certificate. Read-only.
- `IsValid As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the signature was successfully validated following signature verification. Read-only.
- `IsCertificateExpired As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the digital certificate is expired. Read-only.
- `IsCertificateRevoked As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the digital certificate is revoked. Read-only.
- `IsCertificateUntrusted As Boolean  (read-only)`  
  Gets a Boolean value indicating whether the digital certificate used to digitally sign a document comes from a trusted source. Read-only.

## Methods (5)

- `GetSignatureDetail(sigdet As SignatureDetail) As Variant`  
  Displays a specified detail related to a signature.
    - `sigdet As SignatureDetail` (required): An enumerated value specifying which signature detail to display.
- `GetCertificateDetail(certdet As CertificateDetail) As Variant`  
  Displays a specified detail related to a digital certificate.
    - `certdet As CertificateDetail` (required): An enumerated value specifying which certificate detail to display.
- `ShowSignatureCertificate(ParentWindow As IUnknown)`  
  Displays the selected or default digital certificate.
    - `ParentWindow As IUnknown` (required): Contains the handle to the window that contains the Certificate dialog box.
- `SelectSignatureCertificate(ParentWindow As IUnknown)`  
  Displays a dialog box that allows users to select which signature certificate to use for signing a document.
    - `ParentWindow As IUnknown` (required): Contains a handle to the window containing the certificate selection dialog box.
- `SelectCertificateDetailByThumbprint(bstrThumbprint As String)`  
  Displays a dialog box containing information about a digital certificate following verification of the user from a thumbprint.
    - `bstrThumbprint As String` (required): Contains information about the signer identified by the thumbprint.
