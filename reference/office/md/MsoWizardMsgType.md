# MsoWizardMsgType

**Type:** Enumeration  
**Library:** Microsoft Office 16.0 Object Library  

Specifies the context under which a wizard's callback procedure is called. Used as an argument in a callback procedure designed for use with a custom wizard.

## Constants (5)

- `msoWizardMsgLocalStateOn` = 1  
  Not supported.
- `msoWizardMsgLocalStateOff` = 2  
  User clicked the right button in the decision or branch balloon.
- `msoWizardMsgShowHelp` = 3  
  User clicked the left button in the decision or branch balloon.
- `msoWizardMsgSuspending` = 4  
  Passed to the ActivateWizard method if msoWizardActSuspend is specified for the Act argument.
- `msoWizardMsgResuming` = 5  
  Passed to the ActivateWizard method if msoWizardActResume is specified for the Act argument.
