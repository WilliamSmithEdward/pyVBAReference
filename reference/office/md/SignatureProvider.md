# SignatureProvider

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD6A3-0000-0000-C000-000000000046}  

Represents a signature provider add-in.

## Methods (9)

- `GenerateSignatureLineImage(siglnimg As SignatureLineImage, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown) As IPictureDisp`  
  Gets a signature line image.
    - `siglnimg As SignatureLineImage` (required): Contains the name of the signature line graphic.
    - `psigsetup As SignatureSetup` (required): Specifies initial settings of the signature provider add-in.
    - `psiginfo As SignatureInfo` (required): Specifies information about the signature provider add-in.
- `ShowSignatureSetup(ParentWindow As IUnknown, psigsetup As SignatureSetup)`  
  Provides a signature provider add-in the opportunity to display the Signature Setup dialog box to the user.
    - `ParentWindow As IUnknown` (required): Contains the handle to the window containing the Signature Setup dialog box.
    - `psigsetup As SignatureSetup` (required): Specifies initial settings of the signature provider.
- `ShowSigningCeremony(ParentWindow As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo)`  
  Provides a signature provider add-in the opportunity to display the Signature dialog box to users, allowing them to specify their identity and then be authenticated.
    - `ParentWindow As IUnknown` (required): Contains the handle to the window containing the Signature dialog box.
    - `psigsetup As SignatureSetup` (required): Specifies initial settings of the signature provider.
    - `psiginfo As SignatureInfo` (required): Specifies information about the signature provider.
- `SignXmlDsig(QueryContinue As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown)`  
  Used to sign the XMLDSIG template.
    - `QueryContinue As IUnknown` (required): Provides a way to query the host application for permission to continue the verification operation.
    - `psigsetup As SignatureSetup` (required): Specifies configuration information about a signature line.
    - `psiginfo As SignatureInfo` (required): Specifies information captured from the signing ceremony.
    - `XmlDsigStream As IUnknown` (required): Represents a stream of data containing XML, which represents an XMLDSIG object.
- `NotifySignatureAdded(ParentWindow As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo)`  
  Used to display a dialog box informing the user that the signing process has completed and providing additional functionality for the add-in.
    - `ParentWindow As IUnknown` (required): Allows the host application to obtain the handle to the window containing the displayed dialog box.
    - `psigsetup As SignatureSetup` (required): Contains initial settings of the signature provider.
    - `psiginfo As SignatureInfo` (required): Contains information about the signature provider add-in.
- `VerifyXmlDsig(QueryContinue As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown, pcontverres As ContentVerificationResults, pcertverres As CertificateVerificationResults)`  
  Verifies a signature based on the signed state of the document and the legitimacy of the certificate used for signing.
    - `QueryContinue As IUnknown` (required): Provides a way to query the host application for permission to continue the verification operation.
    - `psigsetup As SignatureSetup` (required): Specifies configuration information about a signature line.
    - `psiginfo As SignatureInfo` (required): Specifies information captured from the signing ceremony.
    - `XmlDsigStream As IUnknown` (required): Represents a stream of data containing XML, which represents an XMLDSIG object.
    - `pcontverres As ContentVerificationResults` (required): Specifies the status of the signature verification action.
    - `pcertverres As CertificateVerificationResults` (required): Specifies the status of the signing certificate verification.
- `ShowSignatureDetails(ParentWindow As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown, pcontverres As ContentVerificationResults, pcertverres As CertificateVerificationResults)`  
  Provides a signature provider add-in the opportunity to display details about a signed signature line and display additional stored information such as a secure time-stamp.
    - `ParentWindow As IUnknown` (required): Contains the handle to the window containing the signature details.
    - `psigsetup As SignatureSetup` (required): Specifies initial settings of the signature provider.
    - `psiginfo As SignatureInfo` (required): Specifies information about the signed signature line.
    - `XmlDsigStream As IUnknown` (required): Represents a stream of data or binary large object of XML.
    - `pcontverres As ContentVerificationResults` (required): Contains a value representing the results of verifying the signature content.
    - `pcertverres As CertificateVerificationResults` (required): Contains a value representing the results of verifying the signing certification.
- `GetProviderDetail(sigprovdet As SignatureProviderDetail) As Variant`  
  Queries the signature provider add-in for various details.
    - `sigprovdet As SignatureProviderDetail` (required): Contains an enumerated value representing the type of information to query the add-in for.
- `HashStream(QueryContinue As IUnknown, Stream As IUnknown) As SAFEARRAY(Byte)`  
  Allows a signature provider add-in to create a hash value for the document that you can use to determine if the document contents were tampered with after digital signing.
    - `QueryContinue As IUnknown` (required): Provides a way to query the host application for permission to continue the hashing process.
    - `Stream As IUnknown` (required): Contains the data stream.
