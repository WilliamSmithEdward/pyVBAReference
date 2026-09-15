# LetterContent

**Type:** Class  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209F1-0000-0000-C000-000000000046}  

Represents the elements of a letter created by the Letter Wizard.

**Remarks:** Use the GetLetterContent or CreateLetterContent method to return a LetterContent object. The following example retrieves and displays the letter recipient's name from the active document. The following example uses the CreateLetterContent method to create a new LetterContent object, which is then used with the RunLetterWizard method. The CreateLetterContent method creates a LetterContent object; however, there are numerous required arguments. If you want to set only a few properties, use the New keyword to create a new, stand-alone LetterContent object. The following example creates a LetterContent object, sets some of its properties, and then uses the LetterContent object with the RunLetterWizard method to run the Letter Wizard, using the preset values as the default settings. You can duplicate a LetterContent object by using the Duplicate property. The following example retrieves the letter elements in the active document and makes a duplicate copy. The example assigns the duplicate copy to aLetter and resets the recipient's name and address to empty strings.

## Properties (35)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified LetterContent object.
- `Duplicate As LetterContent  (read-only)`  
  Returns a read-only LetterContent object that represents the contents of a letter created by the Letter Wizard.
- `DateFormat As String  (read/write)`  
  Returns or sets the date for a letter created by the Letter Wizard. Read/write String.
- `IncludeHeaderFooter As Boolean  (read/write)`  
  True if the header and footer from the page design template are included in a letter created by the Letter Wizard. Read/write Boolean. Use the PageDesign property to set the name of the template attached to a document created by the Letter Wizard.
- `PageDesign As String  (read/write)`  
  Returns or sets the name of the template attached to the document created by the Letter Wizard. Read/write String.
- `LetterStyle As WdLetterStyle  (read/write)`  
  Returns or sets the layout of a letter created by the Letter Wizard. Read/write WdLetterStyle.
- `Letterhead As Boolean  (read/write)`  
  True if space is reserved for a preprinted letterhead in a letter created by the Letter Wizard. Read/write Boolean. The LetterheadSize property controls the size of the reserved letterhead space.
- `LetterheadLocation As WdLetterheadLocation  (read/write)`  
  Returns or sets the location of the preprinted letterhead in a letter created by the Letter Wizard. Read/write WdLetterheadLocation.
- `LetterheadSize As Single  (read/write)`  
  Returns or sets the amount of space (in points) to be reserved for a preprinted letterhead in a letter created by the Letter Wizard. Read/write Single.
- `RecipientName As String  (read/write)`  
  Returns or sets the name of the person who'll be receiving the letter created by the Letter Wizard. Read/write String.
- `RecipientAddress As String  (read/write)`  
  Returns or sets the mailing address of the person who'll be receiving the letter created by the Letter Wizard. Read/write String.
- `Salutation As String  (read/write)`  
  Returns or sets the salutation text for a letter created by the Letter Wizard. Read/write String.
- `SalutationType As WdSalutationType  (read/write)`  
  Returns or sets the type of salutation for a letter created by the Letter Wizard. Read/write WdSalutationType.
- `RecipientReference As String  (read/write)`  
  Returns or sets the reference line (for example, "In reply to:") for a letter created by the Letter Wizard. Read/write String.
- `MailingInstructions As String  (read/write)`  
  Returns or sets the mailing instruction text for a letter created by the Letter Wizard (for example, "Certified Mail"). Read/write String.
- `AttentionLine As String  (read/write)`  
  Returns or sets the attention line text for a letter created by the Letter Wizard. Read/write String.
- `Subject As String  (read/write)`  
  Returns or sets the subject text of a letter created by the Letter Wizard. Read/write String.
- `EnclosureNumber As Long  (read/write)`  
  Returns or sets the number of enclosures for a letter created by the Letter Wizard. Read/write String.
- `CCList As String  (read/write)`  
  Returns or sets the carbon copy (CC) recipients for a letter created by the Letter Wizard. Read/write String.
- `ReturnAddress As String  (read/write)`  
  Returns or sets the return address for a letter created with the Letter Wizard. Read/write String.
- `SenderName As String  (read/write)`  
  Returns or sets the name of the person creating a letter with the Letter Wizard. Read/write String.
- `Closing As String  (read/write)`  
  Returns or sets the closing text for a letter created by the Letter Wizard (for example, "Sincerely yours"). Read/write String.
- `SenderCompany As String  (read/write)`  
  Returns or sets the company name of the person creating a letter with the Letter Wizard. Read/write String.
- `SenderJobTitle As String  (read/write)`  
  Returns or sets the job title of the person creating a letter with the Letter Wizard. Read/write String.
- `SenderInitials As String  (read/write)`  
  Returns or sets the initials of the person creating a letter with the Letter Wizard. Read/write String.
- `InfoBlock As Boolean  (read/write)`  
  Associated with the Letter Wizard in Microsoft Word. Not used in the U.S. English version of Word.
- `RecipientCode As String  (read/write)`  
  Returns or sets the recipient code. Read/write String.
- `RecipientGender As WdSalutationGender  (read/write)`  
  Returns or sets the recipient's gender, if known. Not used in the U.S. English version of Microsoft Word. Read/write WdSalutationGender.
- `ReturnAddressShortForm As String  (read/write)`  
  Returns or sets the short form address. Read/write String.
- `SenderCity As String  (read/write)`  
  Returns or sets the sender's city. Not used in the U.S. English version of Microsoft Word. Read/write String.
- `SenderCode As String  (read/write)`  
  Returns or sets the sender code. Not used in the U.S. English version of Microsoft Word. Read/write String.
- `SenderGender As WdSalutationGender  (read/write)`  
  Returns or sets the gender used with the salutation. Not used in the U.S. English version of Microsoft Word. Read/write WdSalutationGender.
- `SenderReference As String  (read/write)`  
  Not used in the U.S. English version of Microsoft Word. Read/write String.
