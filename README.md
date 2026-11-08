# Office 2024: Excel, Word, and PowerPoint

`src/Install-Office.ps1` is a self-contained interactive installer for Office LTSC
2024. Open **PowerShell as administrator** on Windows 10 or Windows 11 and run:

```powershell
irm https://raw.githubusercontent.com/Ehsan-U/office2024-installer/master/src/Install-Office.ps1 | iex
```

No repository checkout or additional script files are needed on the Windows
machine. This project is based on [farag2/Install-Office](https://github.com/farag2/Install-Office).

The menu contains exactly three choices:

```text
1. Excel
2. Word
3. PowerPoint

Enter numbers separated by commas: 1,3
```

Enter `1` for Excel, `2` for Word, `3` for PowerPoint, or any comma-separated
combination. Enter `1,2,3` to install all three. Invalid input is retried; spaces
and duplicate numbers are accepted. Press Ctrl+C to cancel.

To run the local checkout instead:

```powershell
& .\src\Install-Office.ps1
```

The installer downloads Microsoft's Office Deployment Tool, checks its Microsoft
Authenticode signature, and runs `/configure` to download and install the selected
apps directly from Microsoft. It uses `ProPlus2024Volume` / `PerpetualVL2024`,
matches the Windows language with English fallback, and excludes all other suite
apps. It uses 64-bit Office on 64-bit Windows, 32-bit on 32-bit Windows, or matches
an existing Click-to-Run installation's architecture. Setup shows installation
progress, and the script waits for completion and checks the exit code.

A valid Office LTSC 2024 license is required for activation; this script installs
Office and does not activate it. Setup accepts Microsoft's EULA. Existing Office
installations may conflict with this edition; the script does not automatically
uninstall them or force running Office apps to close. Reconfiguring an existing
2024 suite can change which of its apps are installed. Temporary installer files
are cleaned up after setup; Office's own logs remain in Windows' TEMP folder.

This entry point does not run `Configure_Office.ps1` or the upstream region and
registry modifications.

## Development checks

Run the dependency-free tests with Windows PowerShell 5.1 or PowerShell 7:

```powershell
powershell -NoProfile -File .\tests\Install-Office.Tests.ps1
# Or: pwsh -NoProfile -File ./tests/Install-Office.Tests.ps1
```

These check menu parsing, all app combinations, generated XML, and mocked
installation success/failure paths without downloading or installing Office.
A real installation still needs to be tested on Windows with an appropriate license.

## Original advanced download/install workflow

The original scripts below remain available for offline downloads and additional
Office products.

## How-to

[How-to](https://github.com/user-attachments/assets/a817fa3b-d41f-46af-9386-cae939628cdb)

* Download [latest](https://github.com/farag2/Install-Office/releases/latest) archive and expand it
* Open PowerShell console as admin and change execution policy

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
  ```

* Open expanded `Office` folder in PowerShell console and choose which Office to download

* Configure which components to download
  * Branch:
    * ProPlus2024Volume
    * O365ProPlusRetail
  * Channel
    * Current
    * PerpetualVL2024
    * SemiAnnual
  * Components
    * Access
    * OneDrive
    * Outlook
    * Word
    * Excel
    * OneNote
    * Publisher
    * PowerPoint
    * Teams
    * Project Pro 2024
    * Visio Pro 2024

  ```powershell
  .\Download.ps1 -Branch ProPlus2024Volume -Channel PerpetualVL2024 -Components Excel, OneDrive, PowerPoint, Word
  .\Download.ps1 -Branch O365ProPlusRetail -Channel Current -Components Excel, OneDrive, Outlook, PowerPoint, Teams, Word
  ```

  * There won't be any messages in console while Office is being downloaded
  * When Office is downloaded you can find a new `Office` folder in script folder.
* Do not move downloaded `Office` folder to another location as `Install.ps1` script has a link to `setup.exe`
* Run `.\Install.ps1` as admin to install Office you downloaded

## Addendum

`Microsoft Office 2024 & Microsoft 365` support `Windows 10` & `Windows 11` only

## `Configure_Office.ps1` provides the following features

<details>
  <summary>List</summary>

* General
* Remove diagnostics tracking scheduled tasks
* Do not send additional diagnostic and usage data to Microsoft
* Disable LinkedIn features in Office applications
* Turn off the cloud features
* Turn on Touch/Mouse Mode

* Word
  * Do not show the Start screen when application starts
  * Do not open e-mail attachments and other uneditable files in reading view
  * Disable Protected View for files originating from the Internet
  * Disable Protected View for files located in potentially unsafe locations
  * Disable Protected View for Outlook attachments
  * Show the ruler
  * Save AutoRecover information every 3 minutes
  * Enable the "Draw" tab
  * Enable the "Developer" tab
  * Remove Adobe Acrobat Pro DC COM Add-ins

* Excel
  * Do not show the Start screen when application starts
  * Disable Protected View for files originating from the Internet
  * Disable Protected View for files located in potentially unsafe locations
  * Disable Protected View for Outlook attachments
  * Save AutoRecover information every 3 minutes
  * Enable the "Draw" tab
  * Enable the "Developer" tab

</details>

## Links

* [Configure Office](https://config.office.com/deploymentsettings)
* Overview of update channels
  * <https://learn.microsoft.com/en-us/deployoffice/overview-update-channels>
  * <https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run>
* [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* [Deploy Office](https://learn.microsoft.com/en-us/deployoffice/deployment-guide-microsoft-365-apps)
* [GetHelpCmd (ex-SaRA)](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365/admin/miscellaneous/get-help-command-line-overview)
* [OffScrubC2R.vbs 2.19](https://github.com/farag2/Office/tree/master/Office_Uninstall)
