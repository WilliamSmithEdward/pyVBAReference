# SignatureProvider

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD6A3-0000-0000-C000-000000000046}  

## Methods (9)

- `GenerateSignatureLineImage(siglnimg As SignatureLineImage, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown) As IPictureDisp`
- `ShowSignatureSetup(ParentWindow As IUnknown, psigsetup As SignatureSetup)`
- `ShowSigningCeremony(ParentWindow As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo)`
- `SignXmlDsig(QueryContinue As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown)`
- `NotifySignatureAdded(ParentWindow As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo)`
- `VerifyXmlDsig(QueryContinue As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown, pcontverres As ContentVerificationResults, pcertverres As CertificateVerificationResults)`
- `ShowSignatureDetails(ParentWindow As IUnknown, psigsetup As SignatureSetup, psiginfo As SignatureInfo, XmlDsigStream As IUnknown, pcontverres As ContentVerificationResults, pcertverres As CertificateVerificationResults)`
- `GetProviderDetail(sigprovdet As SignatureProviderDetail) As Variant`
- `HashStream(QueryContinue As IUnknown, Stream As IUnknown) As SAFEARRAY(Byte)`
