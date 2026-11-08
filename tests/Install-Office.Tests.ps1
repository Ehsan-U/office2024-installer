# Dependency-free checks. No Office download or installation is performed.
$ErrorActionPreference = 'Stop'
$Path = Join-Path $PSScriptRoot '../src/Install-Office.ps1'
$Tokens = $null
$Errors = $null
$Ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$Tokens, [ref]$Errors)
if ($Errors.Count) { throw ($Errors | Out-String) }
# Load only function definitions, leaving the installer entry point unexecuted.
$Ast.FindAll({ param($Node) $Node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object {
    . ([scriptblock]::Create($_.Extent.Text))
}
function Assert($Condition, $Message) {
    if (-not $Condition) { throw $Message }
}
$Cases = @{
    '1' = 'Excel'; '2' = 'Word'; '3' = 'PowerPoint'
    '1,2' = 'Excel,Word'; '1,3' = 'Excel,PowerPoint'; '2,3' = 'Word,PowerPoint'
    '1,2,3' = 'Excel,Word,PowerPoint'; ' 3, 1, 3 ' = 'PowerPoint,Excel'
}
foreach ($Selection in $Cases.Keys) {
    $Apps = @(ConvertTo-OfficeSelection $Selection)
    Assert (($Apps -join ',') -eq $Cases[$Selection]) "Wrong selection: $Selection"
    foreach ($Architecture in @('32', '64')) {
        $Config = New-OfficeConfiguration -Apps $Apps -Architecture $Architecture
        Assert ($Config.Configuration.Add.OfficeClientEdition -eq $Architecture) 'Wrong architecture'
        Assert ($Config.Configuration.Add.Channel -eq 'PerpetualVL2024') 'Wrong channel'
        Assert ($Config.Configuration.Add.Product.ID -eq 'ProPlus2024Volume') 'Wrong product'
        foreach ($App in @('Excel','Word','PowerPoint','Access','Outlook','OneNote','OneDrive','Teams','Publisher','Lync','Groove')) {
            Assert (($null -eq $Config.SelectSingleNode("//ExcludeApp[@ID='$App']")) -eq ($App -in $Apps)) "Wrong inclusion: $App"
        }
        Assert ($null -eq $Config.SelectSingleNode('//RemoveMSI|//Remove|//Property[@Name="FORCEAPPSHUTDOWN"]')) 'Unexpected destructive setting'
    }
}
foreach ($Selection in @('', ' ', '0', '4', '-1', '01', '1,', ',2', '1,,2', '1 2', '1;2', '1,foo')) {
    $Rejected = $false
    try { ConvertTo-OfficeSelection $Selection | Out-Null } catch { $Rejected = $true }
    Assert $Rejected "Accepted invalid selection: $Selection"
}

# Mock only Windows/network/process boundaries and exercise the real orchestration.
function Assert-OfficeEnvironment {}
function Get-OfficeArchitecture { '64' }
function Read-Host {
    $script:ReadCount++
    if ($script:ReadCount -eq 1) { return '1,,2' }
    return '1,3'
}
function Write-Host { param($Object, $ForegroundColor) $script:Messages += [string]$Object }
function Write-Warning { param($Message) }
function Invoke-WebRequest {
    param($Uri, $OutFile, [switch]$UseBasicParsing)
    Assert ($Uri -eq 'https://officecdn.microsoft.com/pr/wsus/setup.exe') 'Wrong download source'
    $script:WorkDirectory = Split-Path $OutFile
    if ($script:DownloadFails) { throw 'Simulated download failure' }
    Set-Content -LiteralPath $OutFile -Value 'mock deployment tool'
}
function Get-AuthenticodeSignature {
    param($LiteralPath)
    [pscustomobject]@{ Status = $script:SignatureStatus; SignerCertificate = [pscustomobject]@{Subject = $script:Signer} }
}
function Start-Process {
    param($FilePath, $ArgumentList, $WorkingDirectory, [switch]$Wait, [switch]$PassThru)
    $script:Launches++
    Assert ($Wait -and $PassThru) 'Must wait and inspect exit code'
    Assert ($ArgumentList -eq "/configure `"$(Join-Path $WorkingDirectory 'Configuration.xml')`"") 'Wrong arguments'
    [xml]$Config = Get-Content -LiteralPath (Join-Path $WorkingDirectory 'Configuration.xml')
    Assert ($null -eq $Config.SelectSingleNode('//ExcludeApp[@ID="Excel"]')) 'Excel not selected'
    Assert ($null -eq $Config.SelectSingleNode('//ExcludeApp[@ID="PowerPoint"]')) 'PowerPoint not selected'
    Assert ($null -ne $Config.SelectSingleNode('//ExcludeApp[@ID="Word"]')) 'Word unexpectedly selected'
    [pscustomobject]@{ ExitCode = $script:ExitCode }
}
foreach ($Scenario in @('success', 'reboot', 'setup-failure', 'download-failure', 'invalid-signature', 'wrong-publisher')) {
    $script:ReadCount = 0
    $script:Launches = 0
    $script:Messages = @()
    $script:WorkDirectory = $null
    $script:DownloadFails = $Scenario -eq 'download-failure'
    $script:SignatureStatus = if ($Scenario -eq 'invalid-signature') { 'NotSigned' } else { 'Valid' }
    $script:Signer = if ($Scenario -eq 'wrong-publisher') { 'CN=Other, O=Other Corporation, C=US' } else { 'CN=Microsoft Corporation, O=Microsoft Corporation, C=US' }
    $script:ExitCode = switch ($Scenario) { 'reboot' {3010} 'setup-failure' {1603} default {0} }
    $Protocol = [Net.ServicePointManager]::SecurityProtocol
    $Failure = $null
    try { Install-SelectedOffice } catch { $Failure = $_.Exception.Message }
    Assert ($script:ReadCount -eq 2) 'Invalid input was not retried'
    Assert (-not (Test-Path -LiteralPath $script:WorkDirectory)) 'Temporary files were not cleaned'
    Assert ([Net.ServicePointManager]::SecurityProtocol -eq $Protocol) 'TLS preference was not restored'
    if ($Scenario -in @('success', 'reboot')) {
        Assert ($null -eq $Failure) "Unexpected failure: $Failure"
        Assert ($script:Launches -eq 1) 'Setup was not launched once'
        if ($Scenario -eq 'reboot') { Assert (($Messages -join ' ') -match 'Restart Windows') 'Missing restart notice' }
    }
    else {
        Assert ($null -ne $Failure) "Failure was swallowed: $Scenario"
        Assert (-not (($Messages -join ' ') -match 'installation completed')) 'False success message'
        if ($Scenario -eq 'setup-failure') {
            Assert ($Failure -match '1603') 'Exit code missing from failure'
        }
        else { Assert ($script:Launches -eq 0) 'Executed an unverified or missing download' }
    }
}
'PASS: selection, XML, installer orchestration, retry, exit codes, signature/download failures, and cleanup.'
