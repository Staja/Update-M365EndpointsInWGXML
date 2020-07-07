#requires -version 5
<#
.SYNOPSIS
  Fetches Microsoft 365 Endpoint Set data and adds them as Aliases in a WatchGuard Profile Configuration XML
.NOTES
  Version:        2.1
  Author:         Staja
  Date:           2020-07-03
  License:        GPLv3
#>
#-----------------------------------------------------------[Parameters]---------------------------------------------------------
[CmdletBinding()]
Param(
    # Source WatchGuard Profile Configuration
    [Parameter(Mandatory=$true)]
    [ValidateScript({
        if( -Not ($_ | Test-Path) ){
            throw "File does not exist"
        }
        if($_ -notmatch "(\.xml)"){
            throw "The file specified in the path argument must be of type XML"
        }
        return $true
    })]
    [System.IO.FileInfo]$WatchGuardConfigXML,

    # Modified WatchGuard Profile Configuration Destination
    [Parameter(Mandatory=$true)]
    [ValidateScript({
        if($_ -notmatch "(\.xml)"){
            throw "The file specified in the path argument must be of type XML"
        }
        return $true
    })]
    [System.IO.FileInfo]$NewWatchGuardConfigXML,

    # Client Request ID for Microsoft 365 Endpoints API Web Service. Use a new GUID per machine and not per request.
    # By default, will use GUID of first available Network Adapter with GUID but you specify your own GUID
    [Guid]$ClientRequestId = [Guid](Get-WmiObject Win32_NetworkAdapter -Property GUID | Where { $_.GUID -ne $null })[0].GUID,

    # Alias name Prefix for created Microsoft 365 Endpoint Sets (Can't be the same as AliasServicePrefix)
    [string]$AliasPrefix = "M365-EndpointSet-",

    # The Format String for the Microsoft 365 Endpoint Set ID that is suffixed to Alias name
    [string]$IdFormat = "{0:d3}",

    # Alias name Prefix for created Microsoft 365 Endpoint Services (Can't be the same as AliasPrefix, used to find existing Service Area + Port(s) combination aliases)
    [string]$AliasServicePrefix = "M365-Service-",

    # The Format String for the Microsoft 365 Service Area + Port(s) combination Alias name
    # {0} is Service Area
    # {1} is matching Port Definition (below)
    [string]$AliasServiceFormat = "M365-Service-{0}-{1}",

    # Script Log File
    [string]$LogFilePath = (Join-Path -Path $PSScriptRoot -ChildPath "$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).$((Get-Date).ToString('yyyyMMdd')).log")
)
#---------------------------------------------------------[Configuration]--------------------------------------------------------

# Microsoft 365 Endpoints API Web Service (Don't forget to include unique GUID as clientrequestid)
$M365EndpointsAPI = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=$($ClientRequestId)"

# Used for Port Definition Name Lookup
$PortDefinitions = @(
    @{
        "Name" = "HTTP"
        "TcpPorts" = @(80)
        "UdpPorts" = @()
    },
    @{
        "Name" = "HTTPS"
        "TcpPorts" = @(443)
        "UdpPorts" = @()
    },
    @{
        "Name" = "Web"
        "TcpPorts" = @(80,443)
        "UdpPorts" = @()
    },
    @{
        "Name" = "MSA"
        "TcpPorts" = @(587)
        "UdpPorts" = @()
    },
    @{
        "Name" = "IMAP4"
        "TcpPorts" = @(143,993)
        "UdpPorts" = @()
    },
    @{
        "Name" = "StreamedMedia"
        "TcpPorts" = @()
        "UdpPorts" = @(3478,3479,3480,3481)
    },
    @{
        "Name" = "POP3"
        "TcpPorts" = @(995)
        "UdpPorts" = @()
    },
    @{
        "Name" = "SMTP"
        "TcpPorts" = @(25)
        "UdpPorts" = @()
    }
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
# Logging
Function Log-Start {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LogFilePath
    )

    If((Test-Path -Path $LogFilePath)){
        Remove-Item -Path $LogFilePath -Force
    }
    New-Item -Path $LogFilePath -ItemType File
    Add-Content -Path $LogFilePath -Value "***************************************************************************************************"
    Add-Content -Path $LogFilePath -Value (Split-Path -Leaf $PSCommandpath)
    Add-Content -Path $LogFilePath -Value ([DateTime]::Now)
    Add-Content -Path $LogFilePath -Value "***************************************************************************************************"
}
Function Write-LogEntry {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Verbose', 'Information','Warning','Error','DEBU','VERB','INFO','WARN','ERRO')]
        [string]$Severity = 'Information',

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LogFilePath
    )

    switch ($Severity) { 
        'DEBU' { 
            Write-Verbose -Message $Message
            Write-Host "$($Severity): $($Message)" -ForegroundColor Gray
        }
        'VERB' { 
            Write-Verbose -Message $Message
            Write-Host "$($Severity): $($Message)" -ForegroundColor Gray
        }
        'Information' { 
            Write-Verbose -Message $Message
            Write-Host "$($Severity): $($Message)"
        } 
        'INFO' { 
            Write-Verbose -Message $Message
            Write-Host "$($Severity): $($Message)"
        }
        'Verbose' { 
            Write-Verbose -Message $Message 
            Write-Host "$($Severity): $($Message)" -ForegroundColor Gray
        }
        'Warning' { 
            Write-Warning -Message $Message
        } 
        'WARN' { 
            Write-Warning -Message $Message
        } 
        'Error' { 
            Write-Error -Message $Message -ErrorAction SilentlyContinue
            Write-Host "$($Severity): $($Message)" -ForegroundColor Red
        }
        'ERRO' { 
            Write-Error -Message $Message -ErrorAction SilentlyContinue
            Write-Host "$($Severity): $($Message)" -ForegroundColor Red
        }
        default {
            Write-Verbose -Message $Message
        }
    }
    Add-Content -Path $LogFilePath -Value "$($Severity) $([DateTime]::Now): $($Message)"
}

Function ConvertTo-NetMask {
    <#
    .SYNOPSIS
        Convert a mask length to a dotted-decimal subnet mask.
    .DESCRIPTION
        ConvertTo-NetMask returns a subnet mask in dotted decimal format from an integer value ranging between 0 and 32.
    .INPUTS
        System.Int32
    .EXAMPLE
        ConvertTo-NetMask 24
        Returns the dotted-decimal form of the mask, 255.255.255.0, as string.
    #>

    [CmdletBinding()]
    [OutputType([String])]
    param (
        # The number of bits which must be masked.
        [Parameter(Mandatory, Position = 1, ValueFromPipeline)]
        [Alias('Length')]
        [ValidateRange(0, 32)]
        [Byte]$MaskLength
    )

    process {
        ([IPAddress][UInt64][Convert]::ToUInt32(('1' * $MaskLength).PadRight(32, '0'), 2)).IPAddressToString
    }
}

Function Create-WGAddressGroupMembersNodeWithSingleMember { 
    <#
    .SYNOPSIS
        Creates a addr-group-member XML node with a single child member node.
        Typically to be added to a address-group XML node.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $XmlDoc,
        [Parameter(Mandatory)]
        $Type,
        [Parameter(Mandatory)]
        $Value
    )
    # addr-group-member
    $AddressGroupMembersNode = $XmlDoc.CreateElement("addr-group-member")

    # addr-group-member.member
    $AddressGroupMemberNode = $XmlDoc.CreateElement("member")

    # addr-group-member.member.type
    $AddressGroupMemberTypeNode = $XmlDoc.CreateElement("type")
    if($Type -eq "Domain") {
        # addr-group-member.member.type
        $AddressGroupMemberTypeNode.InnerText = 8
        $AddressGroupMemberNode.AppendChild($AddressGroupMemberTypeNode) | Out-Null

        # addr-group-member.member.domain
        $AddressGroupMemberDomainNode = $XmlDoc.CreateElement("domain")
        $AddressGroupMemberDomainNode.InnerText = $Value
        $AddressGroupMemberNode.AppendChild($AddressGroupMemberDomainNode) | Out-Null
    } elseif($Type -eq "HostOrNetwork") {
        $NetworkTuple = $Value.Split("/")
        $IP = $NetworkTuple[0]
        $SubnetLength = $NetworkTuple[1]
        if($IP.Contains(":")) {
            # IPv6
            if($SubnetLength -eq "128") {
                # IPv6 Host

                # address-group.addr-group-member.member.type
                $AddressGroupMemberTypeNode.InnerText = 5
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberTypeNode) | Out-Null

                # address-group.addr-group-member.member.host-ip6-addr
                $AddressGroupMemberIPNode = $XmlDoc.CreateElement("host-ip6-addr")
                $AddressGroupMemberIPNode.InnerText = $IP
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberIPNode) | Out-Null
            } else {
                # IPv6 Network

                # address-group.addr-group-member.member.type
                $AddressGroupMemberTypeNode.InnerText = 6
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberTypeNode) | Out-Null

                # address-group.addr-group-member.member.ip6-network-addr
                $AddressGroupMemberIPNode = $XmlDoc.CreateElement("ip6-network-addr")
                $AddressGroupMemberIPNode.InnerText = $Value
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberIPNode) | Out-Null
            }
        } else {
            # IPv4
            if($SubnetLength -eq "32") {
                # IPv4 Host

                # address-group.addr-group-member.member.type
                $AddressGroupMemberTypeNode.InnerText = 1
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberTypeNode) | Out-Null

                # address-group.addr-group-member.member.host-ip-addr
                $AddressGroupMemberIPNode = $XmlDoc.CreateElement("host-ip-addr")
                $AddressGroupMemberIPNode.InnerText = $IP
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberIPNode) | Out-Null
            } else {
                # IPv4 Network    

                # address-group.addr-group-member.member.type
                $AddressGroupMemberTypeNode.InnerText = 2
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberTypeNode) | Out-Null

                # address-group.addr-group-member.member.ip-network-addr
                $AddressGroupMemberIPNode = $XmlDoc.CreateElement("ip-network-addr")
                $AddressGroupMemberIPNode.InnerText = $IP
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberIPNode) | Out-Null

                # address-group.addr-group-member.member.ip-mask
                $AddressGroupMemberIPMaskNode = $XmlDoc.CreateElement("ip-mask")
                $AddressGroupMemberIPMaskNode.InnerText = ConvertTo-NetMask -MaskLength $SubnetLength
                $AddressGroupMemberNode.AppendChild($AddressGroupMemberIPMaskNode) | Out-Null
            }
        }
    }
    $AddressGroupMembersNode.AppendChild($AddressGroupMemberNode) | Out-Null
    return $AddressGroupMembersNode
}


Function Create-WGAddressGroupNodeWithSingleDomain {
    <#
    .SYNOPSIS
        Creates a address-group XML node with an addr-group-member node contains a single child member node of domain type.
        Typically to be added to the address-group-list XML node of a WatchGuard Profile Configuration.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $XmlDoc,
        [Parameter(Mandatory)]
        $AddressGroupName,
        [Parameter(Mandatory)]
        $Domain
    )

    # address-group
    $AddressGroupNode = $XmlDoc.CreateElement("address-group")

    # address-group.name
    $AddressGroupNameNode = $XmlDoc.CreateElement("name")
    $AddressGroupNameNode.InnerText = $AddressGroupName
    $AddressGroupNode.AppendChild($AddressGroupNameNode) | Out-Null

    # address-group.description
    $AddressGroupDescNode = $XmlDoc.CreateElement("description")
    $AddressGroupNode.AppendChild($AddressGroupDescNode) | Out-Null

    # address-group.property
    $AddressGroupPropNode = $XmlDoc.CreateElement("property")
    $AddressGroupPropNode.InnerText = 16
    $AddressGroupNode.AppendChild($AddressGroupPropNode) | Out-Null

    # address-group.addr-group-member
    $AddressGroupMembersNode = Create-WGAddressGroupMembersNodeWithSingleMember -XmlDoc $XmlDoc -Type "Domain" -Value $Domain

    # Add: address-group
    $AddressGroupNode.AppendChild($AddressGroupMembersNode) | Out-Null
    return $AddressGroupNode
}

Function Create-WGAddressGroupNodeWithHostOrNetwork {
    <#
    .SYNOPSIS
        Creates a address-group XML node with an addr-group-member node contains a single child member node of host or network type.
        Typically to be added to the address-group-list XML node of a WatchGuard Profile Configuration.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $XmlDoc,
        [Parameter(Mandatory)]
        $AddressGroupName,
        [Parameter(Mandatory)]
        $HostOrNetwork
    )
    # address-group
    $AddressGroupNode = $WatchGuardConfig.CreateElement("address-group")

    # address-group.name
    $AddressGroupNameNode = $WatchGuardConfig.CreateElement("name")
    $AddressGroupNameNode.InnerText = $AddressGroupName
    $AddressGroupNode.AppendChild($AddressGroupNameNode) | Out-Null

    # address-group.description
    $AddressGroupDescNode = $WatchGuardConfig.CreateElement("description")
    $AddressGroupNode.AppendChild($AddressGroupDescNode) | Out-Null

    # address-group.property
    $AddressGroupPropNode = $WatchGuardConfig.CreateElement("property")
    $AddressGroupPropNode.InnerText = 16
    $AddressGroupNode.AppendChild($AddressGroupPropNode) | Out-Null

    # address-group.addr-group-member
    $AddressGroupMembersNode = Create-WGAddressGroupMembersNodeWithSingleMember -XmlDoc $XmlDoc -Type "HostOrNetwork" -Value $HostOrNetwork

    # Add: address-group
    $AddressGroupNode.AppendChild($AddressGroupMembersNode) | Out-Null
    
    return $AddressGroupNode
}

Function Create-WGAliasMember {
    <#
    .SYNOPSIS
        Creates a alias-member XML node.
        Typically to be added to the alias-member-list node of a alias XML node.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $XmlDoc,
        $Type = "Normal",
        $User = "Any",
        $Address = "Any",
        $Interface = "Any",
        $AliasName = $null
    )
    # alias-member
    $AliasMemberNode = $XmlDoc.CreateElement("alias-member")
        
    # alias-member.type
    $AliasMemberTypeNode = $XmlDoc.CreateElement("type")
    if($Type -eq "Alias") {
        # alias-member.type
        $AliasMemberTypeNode.InnerText = 2
        $AliasMemberNode.AppendChild($AliasMemberTypeNode) | Out-Null

        # alias-member.alias-name
        $AliasMemberAliasNameNode = $XmlDoc.CreateElement("alias-name")
        $AliasMemberAliasNameNode.InnerText = $AliasName
        $AliasMemberNode.AppendChild($AliasMemberAliasNameNode) | Out-Null
    } else {
        # alias-member.type
        $AliasMemberTypeNode.InnerText = 1
        $AliasMemberNode.AppendChild($AliasMemberTypeNode) | Out-Null

        # alias-member.user
        $AliasMemberUserNode = $XmlDoc.CreateElement("user")
        $AliasMemberUserNode.InnerText = $User
        $AliasMemberNode.AppendChild($AliasMemberUserNode) | Out-Null

        # alias-member.address
        $AliasMemberAddressNode = $XmlDoc.CreateElement("address")
        $AliasMemberAddressNode.InnerText = $Address
        $AliasMemberNode.AppendChild($AliasMemberAddressNode) | Out-Null

        # alias-member.interface
        $AliasMemberInterfaceNode = $XmlDoc.CreateElement("interface")
        $AliasMemberInterfaceNode.InnerText = $Interface
        $AliasMemberNode.AppendChild($AliasMemberInterfaceNode) | Out-Null
    }
    return $AliasMemberNode
}

Function Create-WGAliasWithoutAliasMemberList {
    <#
    .SYNOPSIS
        Creates an alias XML node.
        Typically to be added to the alias-list node of a WatchGuard Profile Configuration
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $XmlDoc,
        [Parameter(Mandatory)]
        $AliasName,
        [Parameter(Mandatory)]
        $AliasDescription
    )

    # alias
    $AliasNode = $XmlDoc.CreateElement("alias")

    # alias.name
    $AliasNameNode = $XmlDoc.CreateElement("name")
    $AliasNameNode.InnerText = $AliasName
    $AliasNode.AppendChild($AliasNameNode) | Out-Null

    # alias.description
    $AliasDescNode = $XmlDoc.CreateElement("description")
    $AliasDescNode.InnerText = $AliasDescription
    $AliasNode.AppendChild($AliasDescNode) | Out-Null

    # alias.property
    $AliasPropNode = $WatchGuardConfig.CreateElement("property")
    $AliasPropNode.InnerText = 0
    $AliasNode.AppendChild($AliasPropNode) | Out-Null

    return $AliasNode
}

Log-Start -LogFilePath $LogFilePath

# Dump Parameters to log
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "LogFilePath: $($LogFilePath)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "M365EndpointsAPI: $($M365EndpointsAPI)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "WatchGuardConfigXML: $($WatchGuardConfigXML)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "NewWatchGuardConfigXML: $($NewWatchGuardConfigXML)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "AliasPrefix: $($AliasPrefix)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "IdFormat: $($IdFormat)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "AliasServicePrefix: $($AliasServicePrefix)"
Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "AliasServiceFormat: $($AliasServiceFormat)"
foreach($PortDefinition in $PortDefinitions) {
    Write-LogEntry -LogFilePath $LogFilePath -Severity VERB -Message "PortDefinition[]: $($PortDefinition["Name"])"
}

#---------------------------------------------------------[Script]--------------------------------------------------------
# Load Microsoft 365 Endpoint Set JSON Data
Write-LogEntry -LogFilePath $LogFilePath -Severity INFO -Message "Fetching M365 Endpoint Set Data from $($M365EndpointsAPI)"
$M365EndpointsData = Invoke-WebRequest -Uri $M365EndpointsAPI | ConvertFrom-Json

# Load the WatchGuard Profile Configuration XML
Write-LogEntry -LogFilePath $LogFilePath -Severity INFO -Message "Loading existing WatchGuard Profile Configuration from $($WatchGuardConfigXML)"
$WatchGuardConfig = [xml](Get-Content $WatchGuardConfigXML)

# Find all, if any, existing M365 Endpoint Set aliases in the WatchGuard Profile Configuration 
$WGAliasM365EndpointSets = $WatchGuardConfig.profile.'alias-list'.alias | Where { $_.name -like "$($AliasPrefix)*" -and $_.name -notlike "$($AliasServicePrefix)*" }
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "Found $($WGAliasM365EndpointSets.Count)x existing WG-Aliases for M365 Endpoint Sets in the WatchGuard Profile Configuration"

# Find all, if any, existing M365 Service Area aliases in the WatchGuard Profile Configuration 
$WGAliasM365ServiceAreas = $WatchGuardConfig.profile.'alias-list'.alias | Where { $_.name -notlike "$($AliasPrefix)*" -and $_.name -like "$($AliasServicePrefix)*" }
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "Found $($WGAliasM365ServiceAreas.Count)x WG-Aliases for M365 Service Area + Port Combo(s) in the WatchGuard Profile Configuration"

# Determine which M365 Endpoint Sets are new and which already exist in the WatchGuard Profile Configuration
$NewM365EndpointSets = $M365EndpointsData | Where { "$($AliasPrefix)$($IdFormat -f $_.id)" -notin $WGAliasM365EndpointSets.name }
$ExistingM365EndpointSets = $M365EndpointsData | Where { "$($AliasPrefix)$($IdFormat -f $_.id)" -in $WGAliasM365EndpointSets.name }
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($NewM365EndpointSets.Count)x of the M365 Endpoint Sets don't have an WG-Alias"
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($ExistingM365EndpointSets.Count)x of the M365 Endpoint Sets already have an WG-Alias"

# Map M365 Endpoint Sets to Service Area + Port(s) Combinations
$M365ServiceAreaEndpointSets = @{}
$UsedWGServiceAreaAliases = @()
foreach($EndpointSet in $M365EndpointsData) {
    # Microsoft has the tcpPorts and udpPorts as strings for some reason
    if($EndpointSet.tcpPorts -eq $null) {
        $TcpPorts = @()
    } elseif($EndpointSet.tcpPorts.Contains(",")) {
        $TcpPorts = $EndpointSet.tcpPorts.Split(",") | Sort-Object
    } elseif($EndpointSet.tcpPorts.Length -gt 0) {
        $TcpPorts = @($EndpointSet.tcpPorts)
    }
    if($EndpointSet.udpPorts -eq $null) {
        $UdpPorts = @()
    } elseif($EndpointSet.udpPorts.Contains(",")) {
        $UdpPorts = $EndpointSet.udpPorts.Split(",") | Sort-Object
    } elseif($EndpointSet.udpPorts.Length -gt 0) {
        $UdpPorts = @($EndpointSet.tcpPorts)
    }

    # Create crappy human-readable Index
    $PortsInfo = ""
    if($TcpPorts.Count -gt 0) {
        $PortsInfo = "TCP=$($TcpPorts -join ",");"
    }
    if($UdpPorts.Count -gt 0) {
         $PortsInfo = "UDP=$($UdpPorts -join ",");"
    }
    $Key = "$($EndpointSet.serviceArea): $($PortsInfo)"

    if($M365ServiceAreaEndpointSets.ContainsKey($Key)) {
        # Service Area + Port(s) combo already defined, add Endpoint Set it
        $M365ServiceAreaEndpointSets[$Key]["EndpointSetIds"] += @($EndpointSet.id)
    } else {
        # Service Area + Port(s) combo not defined, define it and add Endpoint Set to it

        # Find the friendly Port Name for given Service Ports
        $PortName = $null
        foreach($PortDefinition in $PortDefinitions) {
            if((Compare-Object $PortDefinition["TcpPorts"] $TcpPorts -PassThru).Count -eq 0 -and (Compare-Object $PortDefinition["UdpPorts"] $UdpPorts -PassThru).Count -eq 0) {
                $PortName = $PortDefinition["Name"]
                $UsedWGServiceAreaAlias = $AliasServiceFormat -f $EndpointSet.serviceArea, $PortName
                break;
            }
        }

        # Couldn't find friendly Port Name, just find next available numbered increment
        if([string]::IsNullOrEmpty($PortName)) {
            do {
                $UsedWGServiceAreaAlias = $AliasServiceFormat -f $EndpointSet.serviceArea, $i
                if($UsedWGServiceAreaAlias -notin $UsedWGServiceAreaAliases) {
                    $PortName = $i
                    break;
                }
            } while($true) # Dangerous but eh
        }

        $UsedWGServiceAreaAliases += @($UsedWGServiceAreaAlias)
        
        # Define Service Area + Port(s) Combination
        $M365ServiceAreaEndpointSets[$Key] = @{
            "EndpointSetIds" = @($EndpointSet.id) #
            "ServiceArea" = $EndpointSet.serviceArea #
            "ServiceAreaDisplayName" = $EndpointSet.serviceAreaDisplayName #
            "FirstNotes" = $EndpointSet.notes #
            "PortName" = $PortName #
            "TcpPorts" = $TcpPorts
            "UdpPorts" = $UdpPorts
        }
    }
}
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "Derrived $($M365ServiceAreaEndpointSets.Count)x M365 Service Area + Port Combo(s) that will have WG-Aliases"
foreach($UsedWGServiceAreaAlias in $UsedWGServiceAreaAliases) {
    Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  Alias:$($UsedWGServiceAreaAlias)"
}

# Determine which M365 Service Areas are new and which already exist in the WatchGuard Profile Configuration
$NewM365ServiceAreas = $M365ServiceAreaEndpointSets.Keys | Where { ($AliasServiceFormat -f $M365ServiceAreaEndpointSets[$_]["ServiceArea"], $M365ServiceAreaEndpointSets[$_]["PortName"]) -notin $WGAliasM365ServiceAreas.name }
$ExistingM365ServiceAreas = $M365ServiceAreaEndpointSets.Keys | Where { ($AliasServiceFormat -f $M365ServiceAreaEndpointSets[$_]["ServiceArea"], $M365ServiceAreaEndpointSets[$_]["PortName"]) -in $WGAliasM365ServiceAreas.name }
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($NewM365ServiceAreas.Count)x of the M365 Service Area + Port Combo(s) don't have an WG-Alias"
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($ExistingM365ServiceAreas.Count)x of the M365 Service Area + Port Combo(s) already have an WG-Alias"


# Go through all the existing M365 Endpoint sets and update the defined addresses in the existing respective Aliases
foreach($ExistingM365EndpointSet in $ExistingM365EndpointSets) {
    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id): Existing. $($ExistingM365EndpointSet.urls.Count)x URLs, $($ExistingM365EndpointSet.ips.Count)x IPs/networks"

    # Find the relevant Alias for this M365 Endpoint Set
    # .profile.address-group-list.alias-list.alias
    $AliasNode = $WatchGuardConfig.profile.'alias-list'.alias | Where { $_.name -eq "$($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id)" }
    Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  Found WG-Alias:$($AliasNode.name) M365 Endpoint Set"

    # Remove all existing M365 Endpoint Set Address Groups, we'll be recreating them. We're only reusing the aliases
    # Clean Up: address-group-list, alias-member-list
    $AddressGroupNodes = $WatchGuardConfig.profile.'address-group-list'.'address-group' | Where { $_.name -like "$($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id).*" }
    Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  Found $($AddressGroupNodes.Count)x WG-AddressGroups for M365 Endpoint Sets"
    foreach($AddressGroupNode in $AddressGroupNodes) {
        Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  Removing WG-AddressGroup:$($AddressGroupNode.name)"
        $WatchGuardConfig.profile.'address-group-list'.RemoveChild($AddressGroupNode) | Out-Null
    }
    Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  Removing WG-AliasMemberList on WG-Alias:$($AliasNode.name)"
    $AliasNode.RemoveChild($AliasNode.GetElementsByTagName("alias-member-list")[0]) | Out-Null

    # alias.alias-member-list
    $AliasMemberListNode = $WatchGuardConfig.CreateElement("alias-member-list")
    
    $i = 1 # Address Group Counter for this M365 Endpoint Set
    
    # Create a new Address Group for each URL (domain) in the M365 Endpoint Set and add the Address Group to the Alias
    foreach($EndpointURL in $ExistingM365EndpointSet.urls) {
        # address-group
        $AddressGroupNode = Create-WGAddressGroupNodeWithSingleDomain -XmlDoc $WatchGuardConfig -AddressGroupName "$($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id).$($i).alm" -Domain $EndpointURL

        # Add new Address Group to the WatchGuard Profile Configuration
        $WatchGuardConfig.profile.'address-group-list'.AppendChild($AddressGroupNode) | Out-Null
        Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "  $($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id): New WG-AddressGroup for URL:$($EndpointURL)`r`n$($AddressGroupNode.InnerXml)"

        # alias-member
        $AliasMemberNode = Create-WGAliasMember -XmlDoc $WatchGuardConfig -Address $AddressGroupNode.name
        
        # Add new Alias Member to existing Alias in the WatchGuard Profile Configuration
        $AliasMemberListNode.AppendChild($AliasMemberNode) | Out-Null
        $i = $i + 1
    }

    # Create a new Address Group for each IP/network in the M365 Endpoint Set and add the Address Group to the Alias
    foreach($EndpointNetwork in $ExistingM365EndpointSet.ips) {
        # address-group
        $AddressGroupNode = Create-WGAddressGroupNodeWithHostOrNetwork -XmlDoc $WatchGuardConfig -AddressGroupName "$($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id).$($i).alm" -HostOrNetwork $EndpointNetwork
        
        # Add new Address Group to the WatchGuard Profile Configuration
        $WatchGuardConfig.profile.'address-group-list'.AppendChild($AddressGroupNode) | Out-Null
        Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "  $($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id): New WG-AddressGroup for IP:$($EndpointNetwork)`r`n$($AddressGroupNode.InnerXml)"

        # alias-member
        $AliasMemberNode = Create-WGAliasMember -XmlDoc $WatchGuardConfig -Address $AddressGroupNode.name

        # Add new Alias Member to existing Alias in the WatchGuard Profile Configuration
        $AliasMemberListNode.AppendChild($AliasMemberNode) | Out-Null
        $i = $i + 1
    }
    $AliasNode.AppendChild($AliasMemberListNode) | Out-Null

    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id): Updated WG-Alias:$($AliasNode.name)`r`n$($AliasNode.InnerXml)"
}

# Go through all the existing M365 Endpoint sets and create the defined addresses in new respective Aliases
foreach($NewM365EndpointSet in $NewM365EndpointSets) {
    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasPrefix)$($IdFormat -f $NewM365EndpointSet.id): New. $($NewM365EndpointSet.urls.Count)x URLs, $($NewM365EndpointSet.ips.Count)x IPs/networks"

    # alias
    if([String]::IsNullOrEmpty($NewM365EndpointSet.notes)) {
        $AliasDescription = $NewM365EndpointSet.serviceAreaDisplayName
    } else {
        $AliasDescription = $NewM365EndpointSet.notes
    }
    $AliasNode = Create-WGAliasWithoutAliasMemberList -XmlDoc $WatchGuardConfig -AliasName "$($AliasPrefix)$($IdFormat -f $NewM365EndpointSet.id)" -AliasDescription $AliasDescription
    Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  New WG-Alias:$($AliasNode.name) for M365 Endpoint Set"

    # alias.alias-member-list
    $AliasMemberListNode = $WatchGuardConfig.CreateElement("alias-member-list")

    $i = 1 # Address Group Counter for this M365 Endpoint Set

    # Create a new Address Group for each URL (domain) in the M365 Endpoint Set and add the Address Group to the Alias
    foreach($EndpointURL in $NewM365EndpointSet.urls) {
        $AddressGroupNode = Create-WGAddressGroupNodeWithSingleDomain -XmlDoc $WatchGuardConfig -AddressGroupName "$($AliasPrefix)$($IdFormat -f $NewM365EndpointSet.id).$($i).alm" -Domain $EndpointURL
        $WatchGuardConfig.profile.'address-group-list'.AppendChild($AddressGroupNode) | Out-Null
        Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "  $($AliasPrefix)$($IdFormat -f $NewM365EndpointSet.id): New WG-AddressGroup for URL:$($EndpointURL)`r`n$($AddressGroupNode.InnerXml)"

        $AliasMemberNode = Create-WGAliasMember -XmlDoc $WatchGuardConfig -Address $AddressGroupNode.name
        $AliasMemberListNode.AppendChild($AliasMemberNode) | Out-Null

        $i = $i + 1
    }

    # Create a new Address Group for each IP/network in the M365 Endpoint Set and add the Address Group to the Alias
    foreach($EndpointNetwork in $NewM365EndpointSet.ips) {
        # address-group
        $AddressGroupNode = Create-WGAddressGroupNodeWithHostOrNetwork -XmlDoc $WatchGuardConfig -AddressGroupName "$($AliasPrefix)$($IdFormat -f $NewM365EndpointSet.id).$($i).alm" -HostOrNetwork $EndpointNetwork
        
        # Add new Address Group to the WatchGuard Profile Configuration
        $WatchGuardConfig.profile.'address-group-list'.AppendChild($AddressGroupNode) | Out-Null
        Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "  $($AliasPrefix)$($IdFormat -f $ExistingM365EndpointSet.id): New WG-AddressGroup for IP:$($EndpointNetwork)`r`n$($AddressGroupNode.InnerXml)"

        # alias-member
        $AliasMemberNode = Create-WGAliasMember -XmlDoc $WatchGuardConfig -Address $AddressGroupNode.name

        # Add: alias-member
        $AliasMemberListNode.AppendChild($AliasMemberNode) | Out-Null
        $i = $i + 1
    }
    $AliasNode.AppendChild($AliasMemberListNode) | Out-Null

    # Add new Alias to WatchGuard Profile Configuration
    $WatchGuardConfig.profile.'alias-list'.AppendChild($AliasNode) | Out-Null
    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasPrefix)$($IdFormat -f $NewM365EndpointSet.id): New WG-Alias:$($AliasNode.name)`r`n$($AliasNode.InnerXml)"
}


# Update an Alias Member List for each existing M365 Service Area + Port(s) Combination
foreach($M365ServiceAreaKey in $ExistingM365ServiceAreas) {
    $M365ServiceArea = $M365ServiceAreaEndpointSets[$M365ServiceAreaKey]
    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasServiceFormat -f $M365ServiceArea["ServiceArea"], $M365ServiceArea["PortName"]): Existing. $($M365ServiceArea["ServiceAreaDisplayName"]). $($M365ServiceArea["FirstNotes"])"

    # Alias Description
    $AliasDescription = $M365ServiceArea["ServiceAreaDisplayName"]
    if($M365ServiceArea["EndpointSetIds"].Count -eq 1 -and -not [string]::IsNullOrEmpty($M365ServiceArea["FirstNotes"])) {
        # If only one Endpoint Set, and it has notes, then use Notes as Description
        $AliasDescription = $M365ServiceArea["FirstNotes"]
    }

    # Find the relevant Alias for this M365 Service Area + Port(s) combo
    $AliasNode = $WatchGuardConfig.profile.'alias-list'.alias | Where { $_.name -eq ($AliasServiceFormat -f $M365ServiceArea["ServiceArea"], $M365ServiceArea["PortName"]) }

    # Remove the existing M365 Service Area + Port(s) combo Alias Member List on this Alias
    # Clean Up: alias-member-list
    Write-LogEntry -Severity VERB -LogFilePath $LogFilePath -Message "  Removing WG-AliasMemberList on WG-Alias:$($AliasNode.name)"
    $AliasNode.RemoveChild($AliasNode.GetElementsByTagName("alias-member-list")[0]) | Out-Null

    # alias.alias-member-list
    $AliasMemberListNode = $WatchGuardConfig.CreateElement("alias-member-list")

    # Populate Alias Member List with relevant M365 Endpoint Set Aliases
    foreach($M365ServiceAreaEndpointId in $M365ServiceArea["EndpointSetIds"]) {
        # alias-member
        $AliasMemberNode = Create-WGAliasMember -XmlDoc $WatchGuardConfig -Type "Alias" -Alias "$($AliasPrefix)$($IdFormat -f $M365ServiceAreaEndpointId)"

        # Add: alias-member
        $AliasMemberListNode.AppendChild($AliasMemberNode) | Out-Null
    }
    $AliasNode.AppendChild($AliasMemberListNode) | Out-Null

    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasServiceFormat -f $M365ServiceArea["ServiceArea"], $M365ServiceArea["PortName"]): Updated WG-Alias:$($AliasNode.name)`r`n$($AliasNode.InnerXml)"
}

# Create an Alias for each new M365 Service Area + Port(s) Combination
foreach($M365ServiceAreaKey in $NewM365ServiceAreas) {
    $M365ServiceArea = $M365ServiceAreaEndpointSets[$M365ServiceAreaKey]
    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasServiceFormat -f $M365ServiceArea["ServiceArea"], $M365ServiceArea["PortName"]): New. $($M365ServiceArea["ServiceAreaDisplayName"]). $($M365ServiceArea["FirstNotes"])"

    # Alias Description
    $AliasDescription = $M365ServiceArea["ServiceAreaDisplayName"]
    if($M365ServiceArea["EndpointSetIds"].Count -eq 1 -and -not [string]::IsNullOrEmpty($M365ServiceArea["FirstNotes"])) {
        # If only one Endpoint Set, and it has notes, then use Notes as Description
        $AliasDescription = $M365ServiceArea["FirstNotes"]
    }

    # Create Alias node for M365 Service + Port(s) Combo
    $AliasNode = Create-WGAliasWithoutAliasMemberList -XmlDoc $WatchGuardConfig -AliasName ($AliasServiceFormat -f $M365ServiceArea["ServiceArea"], $M365ServiceArea["PortName"]) -AliasDescription $AliasDescription

    # Populate Alias Member List with relevant M365 Endpoint Set Aliases
    $AliasMemberListNode = $WatchGuardConfig.CreateElement("alias-member-list")
    foreach($M365ServiceAreaEndpointId in $M365ServiceArea["EndpointSetIds"]) {
        # alias-member
        $AliasMemberNode = Create-WGAliasMember -XmlDoc $WatchGuardConfig -Type "Alias" -Alias "$($AliasPrefix)$($IdFormat -f $M365ServiceAreaEndpointId)"

        # Add: alias-member
        $AliasMemberListNode.AppendChild($AliasMemberNode) | Out-Null
    }
    $AliasNode.AppendChild($AliasMemberListNode) | Out-Null

    # Add new Alias to WatchGuard Profile Configuration
    $WatchGuardConfig.profile.'alias-list'.AppendChild($AliasNode) | Out-Null
    Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "$($AliasServiceFormat -f $M365ServiceArea["ServiceArea"], $M365ServiceArea["PortName"]): New WG-Alias:$($AliasNode.name)`r`n$($AliasNode.InnerXml)"
}


# Save modified WatchGuard Profile Configuration to new location
Write-LogEntry -Severity INFO -LogFilePath $LogFilePath -Message "Saving modified WatchGuard Profile Configuration XML to: $($NewWatchGuardConfigXML)"
$WatchGuardConfig.Save($NewWatchGuardConfigXML);
