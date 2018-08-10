"# InstallerStuff" 

 Purpose: 
 
  Work-around for a Microsoft bug in Visual Studio Installer projects plugin.
  Incorrect msi file gets generated when you set RemovePreviousVersion=True in 
  the installer project properties. This script can fix it.

 Usage: Call it with a msi file argument.
 Example:
 
  cscript //nologo "Msi_Fix_RemoveExistingProducts_Record.vbs" "SomeCoolInstaller.msi"

 When to use it:
 
  If you want your installer to automatically uninstall older versions of your product.
  In other words, when you want your new version to replace an existing install.
  (BTW When creating an installer for a new revision of your product, change the 
  ProductCode and PackageCode guids but _keep_ the same UpgradeCode.)

 The Microsoft bug:
 
  Microsofts plugin for Visual Studio 2017 Installer projects has a bug that places the 
  uninstall operation too late in the sequence. First it installs the new files, then 
  invokes uninstall for the previous version. This can cause new files to get incorrectly 
  deleted if they have the same name as old files.
  The fix: Do uninstall of an older version _before_ installing new items.
  This function moves the uninstall operation right after InstallValidate, and just before 
  InstallInitialize.

 Tip: 
 
   You can add this as a post build operation in your Visual Studio Installer project. 
   I typically put the call in a bat file and then call the script from there. 
   For example: Select the installer project in solution explorer and view properties (F4)
   Set these two props:
     PostBuildEvent:    "$(ProjectDir)installer_postbuild.bat" "$(BuiltOuputPath)"
     RunPostBuildEvent: On successful build
   
 Additional information:
 
  I hacked this together by looking at some of the sample scripts provided by Microsoft.
  I recommend checking that stuff out. See Windows SDK example scripts...
   https://docs.microsoft.com/en-us/windows/desktop/msi/windows-installer-scripting-examples

