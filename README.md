# Update-M365EndpointsInWGXML

Proof-of-concept script that creates, or updates, aliases for each Microsoft 365 Endpoint Set from [Office 365 IP Address and URL Web Service](https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges) in a specified WatchGuard&#8482; Profile Configuration XML file.

## Prerequisites

This script has only been tested with PowerShell 5 on the Windows platform.

## Installing

Download the script and run it in place. Be aware of your PowerShell signing policy, you may need to Bypass it or a valid publisher may need to sign the script, depending on the security policy of your environment.

## Running

### Overview

This script will attempt to make Aliases for each Endpoint Set and then make useful Aliases for each Service Area and Port(s) Combination which reference the Endpoint Set aliases. It is recommended to use the latter aliases in any actual rules. Although this bloats the alias list, it was done this was to allow for temporary rules using specific Endpoint Sets and self-documentation.

The script will create new aliases if need be or recreate the address lists of existing aliases that match the naming scheme. You should assume any changes to the generated aliases you make will be lost if the script it run again against them.

This script is a proof of concept, sanity checking of the XML to ensure it is a valid WatchGuard Profile Configuration is poor and it assumes a compliant XML structure in a number of places. WatchGuard does not appear to provide an XML Schema Definition for its configuration XML documents.

### Examples

An example use of this script. It is highly recommended to specify your own ClientRequestId.
```
.\Update-M365EndpointsInWGXML.ps1 `
-WatchGuardConfigXML "C:\Users\netadmin\Documents\My WatchGuard\configs\TestRouter.xml" `
-NewWatchGuardConfigXML "C:\Users\netadmin\Documents\My WatchGuard\configs\TestRouter_Experiment.xml" `
```

### Parameters

* **-WatchGuardConfigXML**
  * Source path of the original WatchGuard Profile Configuration XML file you want to create/update M365 Endpoint aliases into.
  * Example: `C:\Users\netadmin\Documents\My WatchGuard\configs\TestRouter.xml`
* **-NewWatchGuardConfigXML**
  * Destination path of the modified WatchGuard Profile Configuration XML file
  * Example: `C:\Users\netadmin\Documents\My WatchGuard\configs\TestRouter_Experiment.xml`
* **-ClientRequestId**
  * The Client GUID as required by the [Office 365 IP Address and URL Web Service](https://docs.microsoft.com/en-us/office365/enterprise/office-365-ip-web-service#common-parameters)
  * It is highly recommended that you specify your own GUID.
  * Default is the next available GUID of a machines network adapter using WMI but you really shouldn't use the default, generate your own using the `New-Guid` PowerShell Cmdlet.
* **-AliasPrefix**
  * Prefix string of all generated Endpoint Set Aliases
  * Cannot be the same as AliasServicePrefix
  * Specify a AliasPrefix that would ensure a wildcard search (Eg: `M365-EndpointSet-*`) does not find aliases with AliasServicePrefix
  * Default: `M365-EndpointSet-`
* **-IdFormat**
  * The format string for the Endpoint Set ID which is the suffix for a endpoint set alias. Mostly used to allow nice names for sorting.
  * `{0}` is the Endpoint Set ID. String is a [.NET format string](https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-numeric-format-strings).
  * Default: `{0:d3}`
* **-AliasServicePrefix**
  * Prefix string of all generated Service Area + Port(s) Combination Aliases
  * Cannot be the same as AliasPrefix
  * Specify a AliasServicePrefix that would ensure a wildcard search (Eg: M365-Service-*) does not find aliases with AliasPrefix
  * Default: `M365-Service-`
* **-AliasServiceFormat**
  * The [.NET format string](https://docs.microsoft.com/en-us/dotnet/standard/base-types/formatting-types) for the generated Microsoft 365 Service Area + Port(s) combination Aliases.
  * `{0}` is Service Area
  * `{1}` is matching Port Definition as defined inside the script
  * Default: `M365-Service-{0}-{1}`
* **-LogFilePath**
  * The log file produced from the script.
  * Default is to place a datestamped log file in the same location as the script (not current working directory)

### Advanced Notes

The script has defined a `$PortDefintions` which is an array of hashtables of friendly names for port combinations. These are suffixed onto generated Microsoft 365 Service Area + Port(s) combination Aliases as per AliasServiceFormat. If the `TcpPorts` and `UdpPorts` of an Endpoint Set don't match a defined Port Defintion, it'll just suffix the next available incremented number.

The exact Web Service API is defined in the script as `$M365EndpointsAPI`, choose from the [available endpoints](https://docs.microsoft.com/en-us/office365/enterprise/office-365-endpoints), depending on what sort of tenant you have but most are probably going to be worldwide.

WatchGuard cannot support domain address-group entries such as `*-files.sharepoint.com` or `autodiscover.*.onmicrosoft.com` so they will be converted to an overly-permissive compatible variant, `*.sharepoint.com` or `*.onmicrosoft.com` respectively, which WatchGuard can accept. It will also check if such a correction exists in address-group already, before adding it again, to prevent dupes.

Script is not very well optimised (especially when appending to arrays) and probably some of the XML operations too.
  
## License

This project is licensed under the GPLv3 License - see the [LICENSE](LICENSE) file for details
