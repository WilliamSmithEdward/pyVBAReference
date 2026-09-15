# WdAlertLevel

**Type:** Enumeration  
**Library:** Microsoft Word 16.0 Object Library  

Specifies the way certain alerts and messages are handled while a macro is running.

## Constants (3)

- `wdAlertsNone` = 0  
  No alerts or message boxes are displayed. If a macro encounters a message box, the default value is chosen and the macro continues.
- `wdAlertsMessageBox` = -2  
  Only message boxes are displayed; errors are trapped and returned to the macro.
- `wdAlertsAll` = -1  
  All message boxes and alerts are displayed; errors are returned to the macro.
