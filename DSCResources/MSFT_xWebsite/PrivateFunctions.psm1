<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-WebSiteByName
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Name
    )

   Get-Website | Where-Object Name -EQ $Name | Write-Output
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Test-WebSiteByName
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Name
    )

   [array]$website = Get-DSCWebSite $Name

   $website.Count | Write-Output
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-WebSiteBinding
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $SiteName
    )

    [PSObject[]] $Bindings
    $Bindings = (Get-ItemProperty -Path (Join-Path "IIS:\Sites\" $SiteName) -Name Bindings).collection

    $CimBindings = foreach ($binding in $bindings)
    {
        $BindingObject = Get-WebBindingObject -BindingInfo $binding
        New-CimInstance -ClassName MSFT_xWebBindingInformation -Namespace root/microsoft/Windows/DesiredStateConfiguration -Property @{Port=[System.UInt16]$BindingObject.Port;
                                                                                                                                       Protocol=$BindingObject.Protocol;
                                                                                                                                       IPAddress=$BindingObject.IPaddress;
                                                                                                                                       HostName=$BindingObject.Hostname;
                                                                                                                                       CertificateThumbprint=$BindingObject.CertificateThumbprint;
                                                                                                                                       CertificateStoreName=$BindingObject.CertificateStoreName} -ClientOnly
    }

    $CimBindings | Write-Output
}

Export-ModuleMember -Function Get-WebSiteByName Test-WebSiteByName, Get-WebSiteBinding