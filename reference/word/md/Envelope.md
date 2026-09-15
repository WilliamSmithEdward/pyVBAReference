# Envelope

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020918-0000-0000-C000-000000000046}  

Represents an envelope attached to a document.

**Remarks:** Use the Envelope property to return the Envelope object. The following example adds an envelope to a new document and sets the distance between the top of the envelope and the address to 2.25 inches. Remarks The Envelope object is available regardless of whether an envelope has been added to the specified document. However, an error occurs if you use one of the following properties when an envelope has not been added to the document: Address, AddressFromLeft, AddressFromTop, FeedSource, ReturnAddress, ReturnAddressFromLeft, ReturnAddressFromTop, and UpdateDocument. The following example demonstrates how to use the On Error GoTo statement to trap the error that occurs if an envelope has not been added to the active document. If, however, an envelope has been added to the document, the recipient address is displayed. Use the Insert method to add an envelope to the specified document. Use the PrintOut method to set the properties of an envelope and print it without adding it to the document.

## Properties (28)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Envelope object.
- `Address As Range  (read-only)`  
  Returns the envelope delivery address as a Range object. Read-only.
- `ReturnAddress As Range  (read-only)`  
  Returns a Range object that represents the envelope return address.
- `DefaultPrintFIMA As Boolean  (read/write)`  
  True to add a Facing Identification Mark (FIM-A) to envelopes by default. Read/write Boolean.
- `DefaultHeight As Single  (read/write)`  
  Returns or sets the default envelope height, in points. Read/write Single.
- `DefaultWidth As Single  (read/write)`  
  Returns or sets the default envelope width, in points. Read/write Single.
- `DefaultSize As String  (read/write)`  
  Returns or sets the default envelope size. Read/write String.
- `DefaultOmitReturnAddress As Boolean  (read/write)`  
  True if the return address is omitted from envelopes by default. Read/write Boolean.
- `FeedSource As WdPaperTray  (read/write)`  
  Returns or sets the paper tray for the envelope. Read/write WdPaperTray.
- `AddressFromLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the envelope and the delivery address. Read/write Single.
- `AddressFromTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the top edge of the envelope and the delivery address. Read/write Single.
- `ReturnAddressFromLeft As Single  (read/write)`  
  Returns or sets the distance (in points) between the left edge of the envelope and the return address. Read/write Single.
- `ReturnAddressFromTop As Single  (read/write)`  
  Returns or sets the distance (in points) between the top edge of the envelope and the return address. Read/write Single.
- `AddressStyle As Style  (read-only)`  
  Returns a Style object that represents the delivery address style for the envelope. Read-only.
- `ReturnAddressStyle As Style  (read-only)`  
  Returns a Style object that represents the return address style for the envelope.
- `DefaultOrientation As WdEnvelopeOrientation  (read/write)`  
  Returns or sets the default orientation for feeding envelopes. Read/write WdEnvelopeOrientation.
- `DefaultFaceUp As Boolean  (read/write)`  
  True if envelopes are fed face up by default. Read/write Boolean.
- `Vertical As Boolean  (read/write)`  
  True vertically orients text on Asian envelopes. Read/write Boolean.
- `RecipientNamefromLeft As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the recipient's name from the left edge of the envelope. Read/write.
- `RecipientNamefromTop As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the recipient's name from the top edge of the envelope. Read/write.
- `RecipientPostalfromLeft As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the recipient's postal code from the left edge of the envelope. Read/write.
- `RecipientPostalfromTop As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the recipient's postal code from the top edge of the envelope. Read/write.
- `SenderNamefromLeft As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the sender's name from the left edge of the envelope. Read/write.
- `SenderNamefromTop As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the sender's name from the top edge of the envelope. Read/write.
- `SenderPostalfromLeft As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the sender's postal code from the left edge of the envelope. Read/write.
- `SenderPostalfromTop As Single  (read/write)`  
  Returns or sets a Single that represents the position, measured in points, of the sender's postal code from the top edge of the envelope. Read/write.

## Methods (4)

- `UpdateDocument()`  
  Updates the envelope in the document with the current envelope settings.
- `Options()`  
  Displays the Envelope Options dialog box.
- `Insert([ExtractAddress As Variant], [Address As Variant], [AutoText As Variant], [OmitReturnAddress As Variant], [ReturnAddress As Variant], [ReturnAutoText As Variant], [PrintBarCode As Variant], [PrintFIMA As Variant], [Size As Variant], [Height As Variant], [Width As Variant], [FeedSource As Variant], [AddressFromLeft As Variant], [AddressFromTop As Variant], [ReturnAddressFromLeft As Variant], [ReturnAddressFromTop As Variant], [DefaultFaceUp As Variant], [DefaultOrientation As Variant], [PrintEPostage As Variant], [Vertical As Variant], [RecipientNamefromLeft As Variant], [RecipientNamefromTop As Variant], [RecipientPostalfromLeft As Variant], [RecipientPostalfromTop As Variant], [SenderNamefromLeft As Variant], [SenderNamefromTop As Variant], [SenderPostalfromLeft As Variant], [SenderPostalfromTop As Variant])`  
  Inserts an envelope as a separate section at the beginning of the specified document.
    - `ExtractAddress As Variant` (optional): True to use the text marked by the EnvelopeAddress bookmark (a user-defined bookmark) as the recipient's address.
    - `Address As Variant` (optional): A string that specifies the recipient's address (ignored if ExtractAddress is True).
    - `AutoText As Variant` (optional): A string that specifies an AutoText entry to use for the address. If specified, Address is ignored.
    - `OmitReturnAddress As Variant` (optional): True to not insert a return address.
    - `ReturnAddress As Variant` (optional): A string that specifies the return address.
    - `ReturnAutoText As Variant` (optional): A string that specifies an AutoText entry to use for the return address. If specified, ReturnAddress is ignored.
    - `PrintBarCode As Variant` (optional): True to add a POSTNET bar code. For U.S. mail only.
    - `PrintFIMA As Variant` (optional): True to add a Facing Identification Mark (FIMA) for use in presorting courtesy reply mail. For U.S. mail only.
    - `Size As Variant` (optional): A string that specifies the envelope size. The string must match one of the sizes listed in the Envelope size box in the Envelope Options dialog box (for example, "Size 10" or "C4").
    - `Height As Variant` (optional): The height of the envelope, measured in points, when the Size argument is set to "Custom size."
    - `Width As Variant` (optional): The width of the envelope, measured in points, when the Size argument is set to "Custom size."
    - `FeedSource As Variant` (optional): True to use the FeedSource property of the Envelope object to specify which paper tray to use when printing the envelope.
    - `AddressFromLeft As Variant` (optional): The distance, measured in points, between the left edge of the envelope and the recipient's address.
    - `AddressFromTop As Variant` (optional): The distance, measured in points, between the top edge of the envelope and the recipient's address.
    - `ReturnAddressFromLeft As Variant` (optional): The distance, measured in points, between the left edge of the envelope and the return address.
    - `ReturnAddressFromTop As Variant` (optional): The distance, measured in points, between the top edge of the envelope and the return address.
    - `DefaultFaceUp As Variant` (optional): True to print the envelope face up, False to print it face down.
    - `DefaultOrientation As Variant` (optional): The orientation for the envelope. Can be any WdEnvelopeOrientation constant.
    - `PrintEPostage As Variant` (optional): True to insert postage from an Internet postage vendor.
    - `Vertical As Variant` (optional): True to print vertical text on the envelope. Used for Asian envelopes. Default is False.
    - `RecipientNamefromLeft As Variant` (optional): Position of the recipient's name, measured in points from the left edge of the envelope. Used for Asian envelopes.
    - `RecipientNamefromTop As Variant` (optional): Position of the recipient's name, measured in points from the top edge of the envelope. Used for Asian envelopes.
    - `RecipientPostalfromLeft As Variant` (optional): Position of the recipient's postal code, measured in points from the left edge of the envelope. Used for Asian envelopes.
    - `RecipientPostalfromTop As Variant` (optional): Position of the recipient's postal code, measured in points from the top edge of the envelope. Used for Asian envelopes.
    - `SenderNamefromLeft As Variant` (optional): Position of the sender's name, measured in points from the left edge of the envelope. Used for Asian envelopes.
    - `SenderNamefromTop As Variant` (optional): Position of the sender's name, measured in points from the top edge of the envelope. Used for Asian envelopes.
    - `SenderPostalfromLeft As Variant` (optional): Position of the sender's postal code, measured in points from the left edge of the envelope. Used for Asian envelopes.
    - `SenderPostalfromTop As Variant` (optional): Position of the sender's postal code, measured in points from the top edge of the envelope. Used for Asian envelopes.
- `PrintOut([ExtractAddress As Variant], [Address As Variant], [AutoText As Variant], [OmitReturnAddress As Variant], [ReturnAddress As Variant], [ReturnAutoText As Variant], [PrintBarCode As Variant], [PrintFIMA As Variant], [Size As Variant], [Height As Variant], [Width As Variant], [FeedSource As Variant], [AddressFromLeft As Variant], [AddressFromTop As Variant], [ReturnAddressFromLeft As Variant], [ReturnAddressFromTop As Variant], [DefaultFaceUp As Variant], [DefaultOrientation As Variant], [PrintEPostage As Variant], [Vertical As Variant], [RecipientNamefromLeft As Variant], [RecipientNamefromTop As Variant], [RecipientPostalfromLeft As Variant], [RecipientPostalfromTop As Variant], [SenderNamefromLeft As Variant], [SenderNamefromTop As Variant], [SenderPostalfromLeft As Variant], [SenderPostalfromTop As Variant])`  
  Prints an envelope without adding the envelope to the active document.
    - `ExtractAddress As Variant` (optional): True to use the text marked by the "EnvelopeAddress" bookmark (a user-defined bookmark) as the recipient's address.
    - `Address As Variant` (optional): A string that specifies the recipient's address (ignored if ExtractAddress is True).
    - `AutoText As Variant` (optional): The name of the AutoText entry that includes a recipient's address.
    - `OmitReturnAddress As Variant` (optional): True to omit the return address.
    - `ReturnAddress As Variant` (optional): A string that specifies the return address.
    - `ReturnAutoText As Variant` (optional): The name of the AutoText entry that includes a return address.
    - `PrintBarCode As Variant` (optional): True to add a POSTNET bar code. For U.S. mail only.
    - `PrintFIMA As Variant` (optional): True to add a Facing Identification Mark (FIM-A) for use in presorting courtesy reply mail. For U.S. mail only.
    - `Size As Variant` (optional): A string that specifies the envelope size. The string should match one of the sizes listed on the left side of the Envelope size box in the Envelope Options dialog box (for example, "Size 10").
    - `Height As Variant` (optional): The height of the envelope (in points) when the Size argument is set to "Custom size".
    - `Width As Variant` (optional): The width of the envelope (in points) when the Size argument is set to "Custom size".
    - `FeedSource As Variant` (optional): True to use the FeedSource property of the Envelope object to specify which paper tray to use when printing the envelope.
    - `AddressFromLeft As Variant` (optional): The distance (in points) between the left edge of the envelope and the recipient's address.
    - `AddressFromTop As Variant` (optional): The distance (in points) between the top edge of the envelope and the recipient's address.
    - `ReturnAddressFromLeft As Variant` (optional): The distance (in points) between the left edge of the envelope and the return address.
    - `ReturnAddressFromTop As Variant` (optional): The distance (in points) between the top edge of the envelope and the return address.
    - `DefaultFaceUp As Variant` (optional): True to print the envelope face up; False to print it face down.
    - `DefaultOrientation As Variant` (optional): The orientation of the envelope. Can be any WdEnvelopeOrientation constant.
    - `PrintEPostage As Variant` (optional): True to print postage using an Internet e-postage vendor.
    - `Vertical As Variant` (optional): True prints text vertically on the envelope. Used for Asian-language envelopes.
    - `RecipientNamefromLeft As Variant` (optional): The position of the recipient's name, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.
    - `RecipientNamefromTop As Variant` (optional): The position of the recipient's name, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.
    - `RecipientPostalfromLeft As Variant` (optional): The position of the recipient's postal code, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.
    - `RecipientPostalfromTop As Variant` (optional): The position of the recipient's postal code, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.
    - `SenderNamefromLeft As Variant` (optional): The position of the sender's name, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.
    - `SenderNamefromTop As Variant` (optional): The position of the sender's name, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.
    - `SenderPostalfromLeft As Variant` (optional): The position of the sender's postal code, measured in points, from the left edge of the envelope. Used for Asian-language envelopes.
    - `SenderPostalfromTop As Variant` (optional): The position of the sender's postal code, measured in points, from the top edge of the envelope. Used for Asian-language envelopes.
