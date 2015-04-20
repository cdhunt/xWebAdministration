
# ValidateWebsite is a helper function used to validate the results 
function ValidateWebsite 
{
    param 
    (
        [object] $Website,

        [string] $Name
    )

    # If a wildCard pattern is not supported by the website provider. 
    # Hence we restrict user to request only one website information in a single request.
    if($Website.Count-gt 1)
    {
        $errorId = "WebsiteDiscoveryFailure"; 
        $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidResult
        $errorMessage = $($LocalizedData.WebsiteDiscoveryFailureError) -f ${Name} 
        $exception = New-Object System.InvalidOperationException $errorMessage 
        $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

        $PSCmdlet.ThrowTerminatingError($errorRecord);
    }
}

# Helper function used to validate website path
function ValidateWebsitePath
{
    param
    (
        [string] $Name,

        [string] $PhysicalPath
    )

    $PathNeedsUpdating = $false

    if((Get-ItemProperty "IIS:\Sites\$Name" -Name physicalPath) -ne $PhysicalPath)
    {
        $PathNeedsUpdating = $true
    }

    $PathNeedsUpdating

}

# Helper function used to validate website bindings
# Returns true if bindings are valid (ie. port, IPAddress & Hostname combinations are unique).

function ValidateWebsiteBindings
{
    Param
    (
        [parameter()]
        [string] 
        $Name,

        [parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $BindingInfo
    )

   
    $Valid = $true

    foreach($binding in $BindingInfo)
    {
        # First ensure that desired binding information is valid ie. No duplicate IPAddres, Port, Host name combinations. 
             
        if (!(EnsurePortIPHostUnique -Port $binding.Port -IPAddress $binding.IPAddress -HostName $Binding.Hostname -BindingInfo $BindingInfo) )
        {
            $errorId = "WebsiteBindingInputInvalidation"; 
            $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidResult
            $errorMessage = $($LocalizedData.WebsiteBindingInputInvalidationError) -f ${Name} 
            $exception = New-Object System.InvalidOperationException $errorMessage 
            $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

            $PSCmdlet.ThrowTerminatingError($errorRecord);
        }
    }     
    
    return compareWebsiteBindings -Name $Name -BindingInfo $BindingInfo
}

function EnsurePortIPHostUnique
{
    param
    (
        [parameter()]
        [System.UInt16] 
        $Port,

        [parameter()]
        [string] 
        $IPAddress,

        [parameter()]
        [string] 
        $HostName,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $BindingInfo,

        [parameter()]
        $UniqueInstances = 0
    )

    foreach ($Binding in $BindingInfo)
    {
        if($binding.Port -eq $Port -and [string]$Binding.IPAddress -eq $IPAddress -and [string]$Binding.HostName -eq $HostName)
        {
            $UniqueInstances += 1
        }
    }

    if($UniqueInstances -gt 1)
    {
        return $false
    }
    else
    {
        return $true
    }
}

# Helper function used to compare website bindings of actual to desired
# Returns true if bindings need to be updated and false if not.
function compareWebsiteBindings
{
    param
    (
        [parameter()]
        [string] 
        $Name,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $BindingInfo
    )
    #Assume bindingsNeedUpdating
    $BindingNeedsUpdating = $false

    #check to see if actual settings have been passed in. If not get them from website
    if($ActualBindings -eq $null)
    {
        $ActualBindings = Get-WebSiteByName $Name | Get-WebBinding

        #Format Binding information: Split BindingInfo into individual Properties (IPAddress:Port:HostName)
        $ActualBindingObjects = @()
        foreach ($ActualBinding in $ActualBindings)
        {
            $ActualBindingObjects += get-WebBindingObject -BindingInfo $ActualBinding
        }
    }
    
    #Compare Actual Binding info ($FormatActualBindingInfo) to Desired($BindingInfo)
    try
    {
        if($BindingInfo.Count -le $ActualBindingObjects.Count)
        {
            foreach($Binding in $BindingInfo)
            {
                $ActualBinding = $ActualBindingObjects | ?{$_.Port -eq $Binding.CimInstanceProperties["Port"].Value}
                if ($ActualBinding -ne $null)
                {
                    if([string]$ActualBinding.Protocol -ne [string]$Binding.CimInstanceProperties["Protocol"].Value)
                    {
                        $BindingNeedsUpdating = $true
                        break
                    }

                    if([string]$ActualBinding.IPAddress -ne [string]$Binding.CimInstanceProperties["IPAddress"].Value)
                    {
                        # Special case where blank IPAddress is saved as "*" in the binding information.
                        if([string]$ActualBinding.IPAddress -eq "*" -AND [string]$Binding.CimInstanceProperties["IPAddress"].Value -eq "") 
                        {
                            #Do nothing
                        }
                        else
                        {
                            $BindingNeedsUpdating = $true
                            break 
                        }                       
                    }

                    if([string]$ActualBinding.HostName -ne [string]$Binding.CimInstanceProperties["HostName"].Value)
                    {
                        $BindingNeedsUpdating = $true
                        break
                    }

                    if([string]$ActualBinding.CertificateThumbprint -ne [string]$Binding.CimInstanceProperties["CertificateThumbprint"].Value)
                    {
                        $BindingNeedsUpdating = $true
                        break
                    }

                    if([string]$ActualBinding.CertificateStoreName -ne [string]$Binding.CimInstanceProperties["CertificateStoreName"].Value)
                    {
                        $BindingNeedsUpdating = $true
                        break
                    }
                }
                else 
                {
                    {
                        $BindingNeedsUpdating = $true
                        break
                    }
                }
            }
        }
        else
        {
            $BindingNeedsUpdating = $true
        }

        $BindingNeedsUpdating

    }
    catch
    {
        $errorId = "WebsiteCompareFailure"; 
        $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidResult
        $errorMessage = $($LocalizedData.WebsiteCompareFailureError) -f ${Name} 
        $errorMessage += $_.Exception.Message
        $exception = New-Object System.InvalidOperationException $errorMessage 
        $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

        $PSCmdlet.ThrowTerminatingError($errorRecord);
    }
}

function UpdateBindings
{
    param
    (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $BindingInfo
    )
    
    #Need to clear the bindings before we can create new ones
    Clear-ItemProperty IIS:\Sites\$Name -Name bindings -ErrorAction Stop

    foreach($binding in $BindingInfo)
    {
        
        $Protocol = $Binding.CimInstanceProperties["Protocol"].Value
        $IPAddress = $Binding.CimInstanceProperties["IPAddress"].Value
        $Port = $Binding.CimInstanceProperties["Port"].Value
        $HostHeader = $Binding.CimInstanceProperties["HostName"].Value
        $CertificateThumbprint = $Binding.CimInstanceProperties["CertificateThumbprint"].Value
        $CertificateStoreName = $Binding.CimInstanceProperties["CertificateStoreName"].Value
                    
        $bindingParams = @{}
        $bindingParams.Add('-Name', $Name)
        $bindingParams.Add('-Port', $Port)
                    
        #Set IP Address parameter
        if($IPAddress -ne $null)
                {
                $bindingParams.Add('-IPAddress', $IPAddress)
            }
        else # Default to any/all IP Addresses
                {
                $bindingParams.Add('-IPAddress', '*')
            }

        #Set protocol parameter
        if($Protocol-ne $null)
                {
                $bindingParams.Add('-Protocol', $Protocol)
            }
        else #Default to Http
                {
                $bindingParams.Add('-Protocol', 'http')
            }

        #Set Host parameter if it exists
        if($HostHeader-ne $null){$bindingParams.Add('-HostHeader', $HostHeader)}

        try
        {
            New-WebBinding @bindingParams -ErrorAction Stop
        }
        Catch
        {
            $errorId = "WebsiteBindingUpdateFailure"; 
            $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidResult
            $errorMessage = $($LocalizedData.WebsiteUpdateFailureError) -f ${Name} 
            $errorMessage += $_.Exception.Message
            $exception = New-Object System.InvalidOperationException $errorMessage 
            $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

            $PSCmdlet.ThrowTerminatingError($errorRecord);
        }

        try
        {
            if($CertificateThumbprint -ne $null)
            {
                $NewWebbinding = get-WebBinding -name $Name -Port $Port
                $newwebbinding.AddSslCertificate($CertificateThumbprint, $CertificateStoreName)
            }
        }
        catch
        {
            $errorId = "WebBindingCertifcateError"; 
            $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidOperation;
            $errorMessage = $($LocalizedData.WebBindingCertifcateError) -f ${CertificateThumbprint} ;
            $errorMessage += $_.Exception.Message
            $exception = New-Object System.InvalidOperationException $errorMessage ;
            $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

            $PSCmdlet.ThrowTerminatingError($errorRecord);
        }
    }
    
}

function get-WebBindingObject
{
    Param
    (
        $BindingInfo
    )

    #First split properties by ']:'. This will get IPv6 address split from port and host name
    $Split = $BindingInfo.BindingInformation.split("[]")
    if($Split.count -gt 1)
    {
        $IPAddress = $Split.item(1)
        $Port = $split.item(2).split(":").item(1)
        $HostName = $split.item(2).split(":").item(2)
    }
    else
    {
        $SplitProps = $BindingInfo.BindingInformation.split(":")
        $IPAddress = $SplitProps.item(0)
        $Port = $SplitProps.item(1)
        $HostName = $SplitProps.item(2)
    }
       
    $WebBindingObject = New-Object PSObject -Property @{Protocol = $BindingInfo.protocol;IPAddress = $IPAddress;Port = $Port;HostName = $HostName;CertificateThumbprint = $BindingInfo.CertificateHash;CertificateStoreName = $BindingInfo.CertificateStoreName}

    return $WebBindingObject
}

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

   [array]$website = Get-WebSiteByName $Name

   $website.Count | Write-Output
}

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

Export-ModuleMember -Function *