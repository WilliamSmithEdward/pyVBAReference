# SignatureInfo

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD6A2-0000-0000-C000-000000000046}  

## Properties (13)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `ReadOnly As Boolean  (read-only)`
- `SignatureProvider As String  (read-only)`
- `SignatureText As String  (read/write)`
- `SignatureImage As IPictureDisp  (read/write)`
- `SignatureComment As String  (read/write)`
- `ContentVerificationResults As ContentVerificationResults  (read-only)`
- `CertificateVerificationResults As CertificateVerificationResults  (read-only)`
- `IsValid As Boolean  (read-only)`
- `IsCertificateExpired As Boolean  (read-only)`
- `IsCertificateRevoked As Boolean  (read-only)`
- `IsCertificateUntrusted As Boolean  (read-only)`

## Methods (5)

- `GetSignatureDetail(sigdet As SignatureDetail) As Variant`
- `GetCertificateDetail(certdet As CertificateDetail) As Variant`
- `ShowSignatureCertificate(ParentWindow As IUnknown)`
- `SelectSignatureCertificate(ParentWindow As IUnknown)`
- `SelectCertificateDetailByThumbprint(bstrThumbprint As String)`
