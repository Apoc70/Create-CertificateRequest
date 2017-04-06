<# 
    .SYNOPSIS 
    This script creates a new certificate request based on an inf file template. Hostnames used are gathered from Exchange virtual directory configurations.

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2017-02-02

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    http://scripts.granikos.eu

    .DESCRIPTION 
    The script queries Exchange Server 2013+ virtual directory hostnames to create a certificate request.
    
    The request is created using an inf file template. You can prepare multiple template files to choose from. 
    Template files are supposed to be stored in the same folder as this PowerShell script.

    The inf file used to create the certificate request is stored on the same directory as the PowerShell script.

    The script queries for the certificate's common name (CN).

    If created, the certificate request is stored in the same directory as the PowerShell script.

    The content of the certificate request file is the CSR to be submitted to a Certificate Authority.
    
    .NOTES 
    Requirements 
    - Windows Server 2012 R2  
    - Exchange Server 2013/2016 Management Shell
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial community release 

    .PARAMETER InfTemplateFile
    Filename of a local .inf file template, default: Default-Template.inf

    .PARAMETER CreateRequest
    Switch to create the certificate request in the local computer certificate store. If not used only a new inf file will be created. 

    .PARAMETER Country
    Certificate DN attribute for Country (C)

    .PARAMETER City
    Certificate DN attribute for Country (S)

    .PARAMETER State
    Certificate DN attribute for Country (L)

    .PARAMETER Organisation
    Certificate DN attribute for Country (O)

    .PARAMETER Department
    Certificate DN attribute for Country (OU)

    .PARAMETER ModernExchangeOnly
    Switch to query Exchange 2013+ only. If not used, all Exchange Servers will be queried.

    .EXAMPLE
    Create a new certificate request inf file used dedicated organizational information. The common name will be determined seperately.
    
    .\Create-CertificateRequest.ps1 -ModernExchangeOnly -Country DE -State NW -City Hueckelhoven -Organisation Varuna -Department IT

    .EXAMPLE
    Create a new certificate request for Exchange 2013+ using the common name only. The common name will be determined seperately.
    
    .\Create-CertificateRequest.ps1 -ModernExchangeOnly -CreateRequest

#>

[CmdletBinding()]
Param(
[string]$Country = '',
[string]$State = '',
[string]$City = '',
[string]$Organisation = '',
[string]$Department = '',
[string]$InfTemplateFile = 'Default-Template.inf',
[switch]$ModernExchangeOnly,
[switch]$CreateRequest
)

# Placeholders used in template file
$CommonNamePlaceholder = '##COMMONNAME##'
$DNSSANPlaceholder = '##DNSSAN##'

# Declare some variables
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$date = Get-Date -Format 'yyyyMMdd'

# Date triggered filenames for individual inf file and certificate request
$RequestInfFileName = "Certificate-Request-$($date).inf"
$RequestFileName = "Certificate-Request-$($date).req"

<#
  Helper function to update progress bar
#>
function Update-ProgressBar {
  [CmdletBinding()]
  param(
    $Status='',
    $Step=1
    )

  $Steps = 8

  Write-Progress -Id 1 -Activity ('Fetching data from {0} Exchange Servers' -f ($script:ExchangeServerCount)) -Status $Status -PercentComplete ((1/$Steps*$Step*100))
}

<#
  Helper function to select common name from gathered Exchange hostnames
#>
function Show-Menu {
  Write-Output 'Exchange Server Hosts'
  for($i=1;$i -le $CertNames.Count;$i++) {
    Write-Host "$i. $($CertNames[$i-1])"
  }

  $selection = Read-Host "Please select Certificate CN (1-$($i-1))"

  return ($selection -1)
}

if($ModernExchangeOnly) {
  # Fetch Exchange 2013+ servers only
  Update-ProgressBar -Status 'Fetching Exchange Server 2013+ servers' -Step 1
  $ExchangeServers = Get-ExchangeServer | Where-Object {$_.IsE15OrLater -eq $true} | Sort-Object -Property Name
}
else {
  Update-ProgressBar -Status 'Fetching Exchange Server 2013+ AND Legacy Exchange servers' -Step 1
  $ExchangeServers = Get-ExchangeServer | Sort-Object -Property Name
}

$script:ExchangeServerCount = ($ExchangeServers | Measure-Object).Count

# Write-Output "Querying the following Exchage Servers"
# $ExchangeServers | Format-Table Name

# Gather virtual directory data

Update-ProgressBar -Status 'Working on AutoDiscoverServiceInternalUri' -Step 1
[array]$CertNames += (Get-ClientAccessServer).AutoDiscoverServiceInternalUri.Host.ToLower() 

Update-ProgressBar -Status 'Working on OutlookAnywhere' -Step 2
$CertNames += ($ExchangeServers | Get-OutlookAnywhere).InternalHostname.HostnameString.ToLower() 
$CertNames += ($ExchangeServers | Get-OutlookAnywhere).ExternalHostname.HostnameString.ToLower()

Update-ProgressBar -Status 'Working on MapiVirtualDirectory (Exchange 2013+)' -Step 3
$CertNames += ($ExchangeServers | Where-Object -FilterScript {$_.AdminDisplayVersion -ilike '*15*'} | Get-MapiVirtualDirectory).InternalUrl.Host 
$CertNames += ($ExchangeServers | Where-Object -FilterScript {$_.AdminDisplayVersion -ilike '*15*'} | Get-MapiVirtualDirectory).ExternalUrl.Host 

Update-ProgressBar -Status 'Working on OabVirtualDirectory' -Step 4
$CertNames += ($ExchangeServers | Get-OabVirtualDirectory).InternalUrl.Host 
$CertNames += ($ExchangeServers | Get-OabVirtualDirectory).ExternalUrl.Host 

Update-ProgressBar -Status 'Working on ActiveSyncVirtualDirectory' -Step 5
$CertNames += ($ExchangeServers | Get-ActiveSyncVirtualDirectory).InternalUrl.Host 
$CertNames += ($ExchangeServers | Get-ActiveSyncVirtualDirectory).ExternalUrl.Host

Update-ProgressBar -Status 'Working on WebServiceVirtualDirectory' -Step 6
$CertNames += ($ExchangeServers | Get-WebServicesVirtualDirectory).InternalUrl.Host 
$CertNames += ($ExchangeServers | Get-WebServicesVirtualDirectory).ExternalUrl.Host 

Update-ProgressBar -Status 'Working on EcpVirtualDirectory' -Step 7
$CertNames += ($ExchangeServers | Get-EcpVirtualDirectory).InternalUrl.Host 
$CertNames += ($ExchangeServers | Get-EcpVirtualDirectory).ExternalUrl.Host 

Update-ProgressBar -Status 'Working on OwaVirtualDirectory' -Step 8
$CertNames += ($ExchangeServers | Get-OwaVirtualDirectory).InternalUrl.Host 
$CertNames += ($ExchangeServers | Get-OwaVirtualDirectory).ExternalUrl.Host 

# Write-Output 'Identified hostnames for Exchange Certificate:'
$CertNames = $CertNames | Select-Object -Unique | Sort-Object

$CNCert = Show-Menu

if (($Country -ne '') -and ($State -ne '') -and ($City -ne '') -and ($Organisation -ne '') -and ($Department -ne '')) {
  # Set common name to full DN
  $CommonName = ('{0},C={1}, S={2}, L={3}, O={4}, OU={5}' -f $($CertNames[$CNCert]), $Country, $State, $City, $Organisation, $Department)
}
else {
  # Set common name to hostname only, as we do not request an EV certificate
  $CommonName = $CertNames[$CNCert] 
}

# Always create SAN entry, even for a single hostname 
foreach($CertName in $CertNames) {
  $DNSSAN += "dns=$($CertName)&"
}

# Remove last & from DNS SAN configuration string
$DNSSAN = ($DNSSAN -replace ".$")

# Import INF template and replace placeholders
$NewRequestInf = (Get-Content -Path $InfTemplateFile).Replace($CommonNamePlaceholder, $CommonName)
$NewRequestInf = $NewRequestInf.Replace($DNSSANPlaceholder, $DNSSAN)

# Save new INF file
$NewRequestInf | Out-File -FilePath (Join-Path -Path $ScriptDir -ChildPath $RequestInfFileName)  -Force 
Write-Output "Certificate request INF file created: $(Join-Path -Path $ScriptDir -ChildPath $RequestInfFileName)"

if($CreateRequest) {
  # Create new certifcate request in local certificate store
  Invoke-Expression -Command "certreq -new $(Join-Path -Path $ScriptDir -ChildPath $RequestInfFileName) $(Join-Path -Path $ScriptDir -ChildPath $RequestFileName)"
  if(!($LastExitCode -eq 0)) {
    throw 'certreq -new command failed'
  }
  else {
    # Ooops
    Write-Output "Certificate request REQ file created: $(Join-Path -Path $ScriptDir -ChildPath $RequestFileName)"
  }
}