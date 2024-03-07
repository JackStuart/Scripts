
# Macros Testing Scripts

These PowerShell scripts are designed to test and verify various registry settings related to macros within Microsoft Office applications. They help ensure that the configured settings align with security best practices from Microsoft and ASD.

## General-Macros.test.ps1

This script validates general macro-related settings across various Office applications as per ASD best practices.

### Settings Checked:
- **Macro Notification Settings**: Require Macros to be signed by a trusted publisher for Access, Excel, Powerpoint, Publisher, Visio, and Word.
- **Block Certificates from Trusted Publishers**: Block certificates from trusted publishers that are only installed in the current user certificate store for Access, Excel, Powerpoint, Publisher, Visio, and Word.
- **Require Extended Key Usage (EKU) for Certificates from Trusted Publishers**: Require Extended Key Usage (EKU) for certificates from trusted publishers for Access, Excel, Powerpoint, Publisher, Visio, and Word.
- **Block Macros from Running in Office Files from the Internet**: Blocking macros from running in Office files from the internet for Access, Excel, Powerpoint, Publisher, Visio, and Word.
- **Allow Trusted Locations on the Network**: Allowing trusted locations on the network for Access, Excel, MS Project, Powerpoint, Visio, and Word.
- **Disable All Trusted Locations**: Disabling all trusted locations for Access, Excel, MS Project, Powerpoint, Visio, and Word.
- **Turn off Trusted Documents on the Network**: Turning off Trusted Documents on the network for Excel, Powerpoint, and Word.
- **Turn off Trusted Documents**: Turning off trusted documents for Excel, Powerpoint, and Word.
- **Trust Access to Visual Basic Project**: Trust access to Visual Basic Project for Excel, Powerpoint, and Word.
- **Allow Mix of Policy and User Locations**: Allowing a mix of policy and user locations for Office applications.
- **Enable Excel 4.0 Macros When VBA Macros are Enabled**: Enabling Excel 4.0 macros when VBA macros are enabled for Excel.
- **Apply Macro Security Settings to Macros, Add-ins, and Additional Actions**: Applying macro security settings to macros, add-ins, and additional actions for Outlook.
- **Enable Microsoft Visual Basic for Applications Project Creation**: Enabling Microsoft Visual Basic for Applications project creation for Visio.
- **Load Microsoft Visual Basic for Applications Projects from Text**: Loading Microsoft Visual Basic for Applications projects from text for Visio.
- 
## Enable-Macros.test.ps1

This script verifies settings that enable macros in Office applications.

### Settings Checked:
- **Macro Notification Settings**: VBA Macro Notification Settings for Access, Excel, MS Project, Powerpoint, Publisher, Visio, and Word.
- **Scan Encrypted Macros in Excel Open XML Workbooks**: Scan encrypted macros in Excel Open XML workbooks.
- **Disable VBA for Office Applications**: Setting to disable VBA for Office applications.
- **Macro Runtime Scan Scope**: Macro Runtime Scan Scope for Office applications.
- **Disable All Trust Bar Notifications for Security Issues**: Disabling Trust Bar notifications for security issues.
- **Automation Security**: Automation Security Level for Office applications.
- **Security Setting for Macros**: Security Level for macros in Outlook.
- **Scan Encrypted Macros in PowerPoint Open XML Workbooks**: Scan encrypted macros in PowerPoint Open XML workbooks.
- **Publisher Automation Security Level**: Automation Security Level for Publisher.
- **Scan Encrypted Macros in Word Open XML Workbooks**: Scan encrypted macros in Word Open XML workbooks.

## Disable-Macros.test.ps1

This script checks for settings that disable macros in Office applications.

### Settings Checked:
- **Macro Notification Settings**: VBA Macro Notification Settings for Access, Excel, MS Project, Powerpoint, Publisher, Visio, and Word.
- **Disable VBA for Office Applications**: Setting to disable VBA for Office applications.
- **Macro Runtime Scan Scope**: Macro Runtime Scan Scope for Office applications.
- **Automation Security**: Automation Security Level for Office applications.
- **Security Setting for Macros**: Security Level for macros in Outlook.
- **Publisher Automation Security Level**: Automation Security Level for Publisher.


## References
https://www.cyber.gov.au/resources-business-and-government/maintaining-devices-and-systems/system-hardening-and-administration/system-hardening/restricting-microsoft-office-macros
