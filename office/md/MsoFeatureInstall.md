# MsoFeatureInstall

**Type:** Enumeration  
**Library:** Microsoft Office 16.0 Object Library  

Specifies how the application handles calls to methods and properties that require features not yet installed.

## Constants (3)

- `msoFeatureInstallNone` = 0  
  Generates a generic automation error at run time when uninstalled features are called.
- `msoFeatureInstallOnDemand` = 1  
  Prompts the user to install new features.
- `msoFeatureInstallOnDemandWithUI` = 2  
  Displays a progress meter during installation; does not prompt the user to install new features.
