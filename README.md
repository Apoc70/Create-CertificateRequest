# Create-CertificateRequest.ps1

This script creates a new certificate request based on an inf file template. Hostnames used are gathered from Exchange virtual directory configurations.

## Description

The script queries Exchange Server 2013+ virtual directory hostnames to create a certificate request.

The request is created using an inf file template. You can prepare multiple template files to choose from.
Template files are supposed to be stored in the same folder as this PowerShell script.

The inf file used to create the certificate request is stored on the same directory as the PowerShell script.

The script queries for the certificate's common name (CN).

If created, the certificate request is stored in the same directory as the PowerShell script.

The content of the certificate request file is the CSR to be submitted to a Certificate Authority.

## Requirements

- Windows Server 2012 R2
- Exchange Server 2013/2016 Management Shell

## Parameters

### InfTemplateFile

Filename of a local .inf file template, default: Default-Template.inf

### CreateRequest

Switch to create the certificate request in the local computer certificate store. If not used only a new inf file will be created.

### Country

Certificate DN attribute for Country (C)

### City

Certificate DN attribute for Country (S)

### State

Certificate DN attribute for Country (L)

### Organisation

Certificate DN attribute for Country (O)

### Department

Certificate DN attribute for Country (OU)

### ModernExchangeOnly

Switch to query Exchange 2013+ only. If not used, all Exchange Servers will be queried.

## Outputs

- The script creates an inf file which can be used as input for certreq.exe
- If selected, the script creates a certificate request using the local computers certificate store and save the certificate request (CSR) to disk

## Examples

``` PowerShell
.\Create-CertificateRequest.ps1 -ModernExchangeOnly -Country DE -State NW -City Hueckelhoven -Organisation Varuna -Department IT
```

Create a new certificate request inf file used dedicated organizational information. The common name will be determined seperately.

``` PowerShell
.\Create-CertificateRequest.ps1 -ModernExchangeOnly -CreateRequest
```

Create a new certificate request for Exchange 2013+ using the common name only. The common name will be determined seperately.

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)