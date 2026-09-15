# AcSysCmdAction

**Type:** Enumeration  
**Library:** Microsoft Access 16.0 Object Library  

Used with the SysCmd method to specify an action to take.

## Constants (21)

- `acSysCmdInitMeter` = 1  
  Initializes the progress meter. You must specify the argument1 and argument2 arguments when you use this action.
- `acSysCmdUpdateMeter` = 2  
  Updates the progress meter with the specified value. You must specify the text argument when you use this action.
- `acSysCmdRemoveMeter` = 3  
  Removes the progress meter.
- `acSysCmdSetStatus` = 4  
  Sets the status bar text to the text argument.
- `acSysCmdClearStatus` = 5  
  Provides information on the state of a database object.
- `acSysCmdRuntime` = 6  
  Returns True (1) if a run-time version of Microsoft Access is running.
- `acSysCmdAccessVer` = 7  
  Returns the version number of Microsoft Access.
- `acSysCmdIniFile` = 8  
  Returns the name of the .ini file associated with Microsoft Access.
- `acSysCmdAccessDir` = 9  
  Returns the name of the directory where Msaccess.exe is located.
- `acSysCmdGetObjectState` = 10  
  Returns the state of the specified database object. You must specify argument1 and argument2 when you use this action value.
- `acSysCmdClearHelpTopic` = 11  
  Resets default help topic.
- `acSysCmdProfile` = 12  
  Returns the \/profile setting specified by the user when starting Microsoft Access from the command line.
- `acSysCmdGetWorkgroupFile` = 13  
  Returns the path to the workgroup file (System.mdw).
- `acSysCmdCompile` = 603  
  Compiles the Visual Basic code modules in the current database. Equivalent to the Debug > Compile menu command.
- `acSysCmdGetMsoBuildNumber` = 715  
  Returns the build number of the shared MSO component as a Long. This is the same value returned by Application.Build, and it may differ from the Access application build. Use acSysCmdGetBuildNumber (725) in new code to get the Access build number.
- `acSysCmdGetFullVersion` = 720  
  Returns a display string containing version, build, channel, and bitness (for example, "Microsoft Access (Version 2601) Build 16.0.19628.20000 Current Channel 64-bit"). Version 2604 and later.
- `acSysCmdGetVersion` = 721  
  Returns the short YYMM marketing version (for example, "2601"). Version 2604 and later.
- `acSysCmdGetFullBuildNumber` = 722  
  Returns the full four-part build string (for example, "16.0.19916.30000"). Version 2604 and later.
- `acSysCmdGetChannelName` = 723  
  Returns the update channel name (for example, "Current Channel", "Monthly Enterprise Channel", or "LTSC 2024"). Version 2604 and later.
- `acSysCmdGetBitness` = 724  
  Returns "32-bit" or "64-bit" as a string matching the bitness of the running binary. Version 2604 and later.
- `acSysCmdGetBuildNumber` = 725  
  Returns the major build number (for example, 19916) as a Long. Version 2604 and later.
