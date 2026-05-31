# ILicAgent

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {00194002-D9C3-11D3-8D59-0050048384E3}  

ILicAgent Interface

## Methods (99)

- `Initialize(dwBPC As Long, dwMode As Long, bstrLicSource As String) As Long`  
  method Initialize
- `GetFirstName() As String`  
  method GetFirstName
- `SetFirstName(bstrNewVal As String)`  
  method SetFirstName
- `GetLastName() As String`  
  method GetLastName
- `SetLastName(bstrNewVal As String)`  
  method SetLastName
- `GetOrgName() As String`  
  method GetOrgName
- `SetOrgName(bstrNewVal As String)`  
  method SetOrgName
- `GetEmail() As String`  
  method GetEmail
- `SetEmail(bstrNewVal As String)`  
  method SetEmail
- `GetPhone() As String`  
  method GetPhone
- `SetPhone(bstrNewVal As String)`  
  method SetPhone
- `GetAddress1() As String`  
  method GetAddress1
- `SetAddress1(bstrNewVal As String)`  
  method SetAddress1
- `GetCity() As String`  
  method GetCity
- `SetCity(bstrNewVal As String)`  
  method SetCity
- `GetState() As String`  
  method GetState
- `SetState(bstrNewVal As String)`  
  method SetState
- `GetCountryCode() As String`  
  method GetCountryCode
- `SetCountryCode(bstrNewVal As String)`  
  method SetCountryCode
- `GetCountryDesc() As String`  
  method GetCountryDesc
- `SetCountryDesc(bstrNewVal As String)`  
  method SetCountryDesc
- `GetZip() As String`  
  method GetZip
- `SetZip(bstrNewVal As String)`  
  method SetZip
- `GetIsoLanguage() As Long`  
  method GetIsoLanguage
- `SetIsoLanguage(dwNewVal As Long)`  
  method SetIsoLanguage
- `GetMSUpdate() As String`  
  method GetMSUpdate
- `SetMSUpdate(bstrNewVal As String)`  
  method SetMSUpdate
- `GetMSOffer() As String`  
  method GetMSOffer
- `SetMSOffer(bstrNewVal As String)`  
  method SetMSOffer
- `GetOtherOffer() As String`  
  method GetOtherOffer
- `SetOtherOffer(bstrNewVal As String)`  
  method SetOtherOffer
- `GetAddress2() As String`  
  method GetAddress2
- `SetAddress2(bstrNewVal As String)`  
  method SetAddress2
- `CheckSystemClock() As Long`  
  method CheckSystemClock
- `GetExistingExpiryDate() As Date`  
  method GetExistingExpiryDate
- `GetNewExpiryDate() As Date`  
  method GetNewExpiryDate
- `GetBillingFirstName() As String`  
  method GetBillingFirstName
- `SetBillingFirstName(bstrNewVal As String)`  
  method SetBillingFirstName
- `GetBillingLastName() As String`  
  method GetBillingLastName
- `SetBillingLastName(bstrNewVal As String)`  
  method SetBillingLastName
- `GetBillingPhone() As String`  
  method GetBillingPhone
- `SetBillingPhone(bstrNewVal As String)`  
  method SetBillingPhone
- `GetBillingAddress1() As String`  
  method GetBillingAddress1
- `SetBillingAddress1(bstrNewVal As String)`  
  method SetBillingAddress1
- `GetBillingAddress2() As String`  
  method GetBillingAddress2
- `SetBillingAddress2(bstrNewVal As String)`  
  method SetBillingAddress2
- `GetBillingCity() As String`  
  method GetBillingCity
- `SetBillingCity(bstrNewVal As String)`  
  method SetBillingCity
- `GetBillingState() As String`  
  method GetBillingState
- `SetBillingState(bstrNewVal As String)`  
  method SetBillingState
- `GetBillingCountryCode() As String`  
  method GetBillingCountryCode
- `SetBillingCountryCode(bstrNewVal As String)`  
  method SetBillingCountryCode
- `GetBillingZip() As String`  
  method GetBillingZip
- `SetBillingZip(bstrNewVal As String)`  
  method SetBillingZip
- `SaveBillingInfo(bSave As Long) As Long`  
  method SaveBillingInfo
- `IsCCRenewalCountry(bstrCountryCode As String) As Long`  
  method IsCCRenewalCountry
- `GetVATLabel(bstrCountryCode As String) As String`  
  method GetVATLabel
- `GetCCRenewalExpiryDate() As Date`  
  method GetCCRenewalExpiryDate
- `SetVATNumber(bstrVATNumber As String)`  
  method SetVATNumber
- `SetCreditCardType(bstrCCCode As String)`  
  method SetCreditCardType
- `SetCreditCardNumber(bstrCCNumber As String)`  
  method SetCreditCardNumber
- `SetCreditCardExpiryYear(dwCCYear As Long)`  
  method SetCreditCardExpiryYear
- `SetCreditCardExpiryMonth(dwCCMonth As Long)`  
  method SetCreditCardExpiryMonth
- `GetCreditCardCount() As Long`  
  method GetCreditCardCount
- `GetCreditCardCode(dwIndex As Long) As String`  
  method GetCreditCardCode
- `GetCreditCardName(dwIndex As Long) As String`  
  method GetCreditCardName
- `GetVATNumber() As String`  
  method GetVATNumber
- `GetCreditCardType() As String`  
  method GetCreditCardType
- `GetCreditCardNumber() As String`  
  method GetCreditCardNumber
- `GetCreditCardExpiryYear() As Long`  
  method GetCreditCardExpiryYear
- `GetCreditCardExpiryMonth() As Long`  
  method GetCreditCardExpiryMonth
- `GetDisconnectOption() As Long`  
  method GetDisconnectOption
- `SetDisconnectOption(bNewVal As Long)`  
  method SetDisconnectOption
- `AsyncProcessHandshakeRequest(bReviseCustInfo As Long)`  
  method AsyncProcessHandshakeRequest
- `AsyncProcessNewLicenseRequest()`  
  method AsyncProcessNewLicenseRequest
- `AsyncProcessReissueLicenseRequest()`  
  method AsyncProcessReissueLicenseRequest
- `AsyncProcessRetailRenewalLicenseRequest()`  
  method AsyncProcessRetailRenewalLicenseRequest
- `AsyncProcessReviseCustInfoRequest()`  
  method AsyncProcessReviseCustInfoRequest
- `AsyncProcessCCRenewalPriceRequest()`  
  method AsyncProcessCCRenewalPriceRequest
- `AsyncProcessCCRenewalLicenseRequest()`  
  method AsyncProcessCCRenewalLicenseRequest
- `GetAsyncProcessReturnCode() As Long`  
  method GetAsyncProcessReturnCode
- `IsUpgradeAvailable() As Long`  
  method IsUpgradeAvailable
- `WantUpgrade(bWantUpgrade As Long)`  
  method WantUpgrade
- `AsyncProcessDroppedLicenseRequest()`  
  method AsyncProcessDroppedLicenseRequest
- `GenerateInstallationId() As String`  
  method GenerateInstallationId
- `DepositConfirmationId(bstrVal As String) As Long`  
  method DepositConfirmationId
- `VerifyCheckDigits(bstrCIDIID As String) As Long`  
  method VerifyCheckDigits
- `GetCurrentExpiryDate() As Date`  
  method GetCurrentExpiryDate
- `CancelAsyncProcessRequest(bIsLicenseRequest As Long)`  
  method CancelAsyncProcessRequest
- `GetCurrencyDescription(dwCurrencyIndex As Long) As String`  
  method GetCurrencyDescription
- `GetPriceItemCount() As Long`  
  method GetPriceItemCount
- `GetPriceItemLabel(dwIndex As Long) As String`  
  method GetPriceItemLabel
- `GetPriceItemValue(dwCurrencyIndex As Long, dwIndex As Long) As String`  
  method GetPriceItemValue
- `GetInvoiceText() As String`  
  method GetInvoiceText
- `GetBackendErrorMsg() As String`  
  method GetBackendErrorMsg
- `GetCurrencyOption() As Long`  
  method GetCurrencyOption
- `SetCurrencyOption(dwCurrencyOption As Long)`  
  method SetCurrencyOption
- `GetEndOfLifeHtmlText() As String`  
  method GetEndOfLifeHtmlText
- `DisplaySSLCert() As Long`  
  method DisplaySSLCert
