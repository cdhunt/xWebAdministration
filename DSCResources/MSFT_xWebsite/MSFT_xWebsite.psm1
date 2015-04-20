Import-Module $PSScriptRoot\Website_HelperFunctions.psm1

data LocalizedData
{
    # culture="en-US"
    ConvertFrom-StringData @'
SetTargetResourceInstallwhatIfMessage=Trying to create website "{0}".
SetTargetResourceUnInstallwhatIfMessage=Trying to remove website "{0}".
WebsiteNotFoundError=The requested website "{0}" is not found on the target machine.
WebsiteDiscoveryFailureError=Failure to get the requested website "{0}" information from the target machine.
WebsiteCreationFailureError=Failure to successfully create the website "{0}".
WebsiteRemovalFailureError=Failure to successfully remove the website "{0}".
WebsiteUpdateFailureError=Failure to successfully update the properties for website "{0}".
WebsiteBindingUpdateFailureError=Failure to successfully update the bindings for website "{0}".
WebsiteBindingInputInvalidationError=Desired website bindings not valid for website "{0}".
WebsiteCompareFailureError=Failure to successfully compare properties for website "{0}".
WebBindingCertifcateError=Failure to add certificate to web binding. Please make sure that the certificate thumbprint "{0}" is valid.
WebsiteStateFailureError=Failure to successfully set the state of the website {0}.
WebsiteBindingConflictOnStartError = Website "{0}" could not be started due to binding conflict. Ensure that the binding information for this website does not conflict with any existing website's bindings before trying to start it.
'@
}

# The Get-TargetResource cmdlet is used to fetch the status of role or Website on the target machine.
# It gives the Website info of the requested role/feature on the target machine.  
function Get-TargetResource 
{
    [OutputType([System.Collections.Hashtable])]
    param 
    (   
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

        $getTargetResourceResult = $null;

        # Check if WebAdministration module is present for IIS cmdlets
        if(!(Get-Module -ListAvailable -Name WebAdministration))
        {
            Throw "Please ensure that WebAdministration module is installed."
        }

        $count = Test-WebSiteByName $Name

        switch ($count)
        {
            0 {
                $ensureResult = "Absent"
                break
            }
            1 {
                $ensureResult = "Present"
                $CimBindings = Get-WebSiteBinding $Name

                $Website = Get-WebSiteByName $Name
                $allDefaultPage = @(Get-WebConfiguration //defaultDocument/files/* -PSPath (Join-Path "IIS:\sites\" $Name) | foreach {Write-Output $_.value})

                break
            }
            {$_ -gt 1} {
                $errorId = "WebsiteDiscoveryFailure"; 
                $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidResult
                $errorMessage = $($LocalizedData.WebsiteUpdateFailureError) -f ${Name} 
                $exception = New-Object System.InvalidOperationException $errorMessage 
                $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

                $PSCmdlet.ThrowTerminatingError($errorRecord);
            }
        }

        
        # Add all Website properties to the hash table
        $getTargetResourceResult = @{
                                        Name = $Website.Name; 
                                        Ensure = $ensureResult;
                                        PhysicalPath = $Website.physicalPath;
                                        State = $Website.state;
                                        ID = $Website.id;
                                        ApplicationPool = $Website.applicationPool;
                                        BindingInfo = $CimBindings;
                                        DefaultPage = $allDefaultPage
                                    }
        
        return $getTargetResourceResult;
}


# The Set-TargetResource cmdlet is used to create, delete or configuure a website on the target machine. 
function Set-TargetResource 
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param 
    (       
        [ValidateSet("Present", "Absent")]
        [string]$Ensure = "Present",

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PhysicalPath,

        [ValidateSet("Started", "Stopped")]
        [string]$State = "Started",

        [string]$ApplicationPool,

        [Microsoft.Management.Infrastructure.CimInstance[]]$BindingInfo,

      [string[]]$DefaultPage

    )
 
    $getTargetResourceResult = $null;

    if($Ensure -eq "Present")
    {
        #Remove Ensure from parameters as it is not needed to create new website
        $Result = $psboundparameters.Remove("Ensure");
        #Remove State parameter form website. Will start the website after configuration is complete
        $Result = $psboundparameters.Remove("State");

        #Remove bindings from parameters if they exist
        #Bindings will be added to site using separate cmdlet
        $Result = $psboundparameters.Remove("BindingInfo");

        #Remove default pages from parameters if they exist
        #Default Pages will be added to site using separate cmdlet
        $Result = $psboundparameters.Remove("DefaultPage");

        # Check if WebAdministration module is present for IIS cmdlets
        if(!(Get-Module -ListAvailable -Name WebAdministration))
        {
            Throw "Please ensure that WebAdministration module is installed."
        }
        $website = Get-WebSiteByName $Name

        if($website -ne $null)
        {
            #update parameters as required

            $UpdateNotRequired = $true

            #Update Physical Path if required
            if(ValidateWebsitePath -Name $Name -PhysicalPath $PhysicalPath)
            {
                $UpdateNotRequired = $false
                Set-ItemProperty "IIS:\Sites\$Name" -Name physicalPath -Value $PhysicalPath -ErrorAction Stop

                Write-Verbose("Physical path for website $Name has been updated to $PhysicalPath");
            }

            #Update Bindings if required
            if ($BindingInfo -ne $null)
            {
                if(ValidateWebsiteBindings -Name $Name -BindingInfo $BindingInfo)
                {
                    $UpdateNotRequired = $false
                    #Update Bindings
                    UpdateBindings -Name $Name -BindingInfo $BindingInfo -ErrorAction Stop

                    Write-Verbose("Bindings for website $Name have been updated.");
                }
            }

            #Update Application Pool if required
            if(($website.applicationPool -ne $ApplicationPool) -and ($ApplicationPool -ne ""))
            {
                $UpdateNotRequired = $false
                Set-ItemProperty IIS:\Sites\$Name -Name applicationPool -Value $ApplicationPool -ErrorAction Stop

                Write-Verbose("Application Pool for website $Name has been updated to $ApplicationPool")
            }

        #Update Default pages if required 
        if($DefaultPage -ne $null)
            {
            UpdateDefaultPages -Name $Name -DefaultPage $DefaultPage 
        }

            #Update State if required
            if($website.state -ne $State -and $State -ne "")
            {
                $UpdateNotRequired = $false
                if($State -eq "Started")
                {
                    # Ensure that there are no other websites with binding information that will conflict with this site before starting
                    $existingSites = Get-Website | Where Name -ne $Name

                    foreach($site in $existingSites)
                    {
                        $siteInfo = Get-TargetResource -Name $site.name
                            
                        foreach ($binding in $BindingInfo)
                        {
                            #Normalize empty IPAddress to "*"
                            if($binding.IPAddress -eq "" -or $binding.IPAddress -eq $null)
                            {
                                $NormalizedIPAddress = "*"
                            } 
                            else
                            {
                                $NormalizedIPAddress = $binding.IPAddress
                            }

                            if( !(EnsurePortIPHostUnique -Port $Binding.Port -IPAddress $NormalizedIPAddress -HostName $binding.HostName -BindingInfo $siteInfo.BindingInfo -UniqueInstances 1))
                            {
                                #return error & Do not start Website
                                $errorId = "WebsiteBindingConflictOnStart";
                                $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidResult
                                $errorMessage = $($LocalizedData.WebsiteBindingConflictOnStartError) -f ${Name} 
                                $exception = New-Object System.InvalidOperationException $errorMessage 
                                $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

                                $PSCmdlet.ThrowTerminatingError($errorRecord);
                            } 
                        }
                    }

                    try
                    {

                    Start-Website -Name $Name

                    }
                    catch
                    {
                        $errorId = "WebsiteStateFailure"; 
                        $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidOperation;
                        $errorMessage = $($LocalizedData.WebsiteStateFailureError) -f ${Name} ;
                        $errorMessage += $_.Exception.Message
                        $exception = New-Object System.InvalidOperationException $errorMessage ;
                        $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

                        $PSCmdlet.ThrowTerminatingError($errorRecord);
                    }
                    
                }
                else
                {
                    try
                    {

                    Stop-Website -Name $Name

                    }
                    catch
                    {
                        $errorId = "WebsiteStateFailure"; 
                        $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidOperation;
                        $errorMessage = $($LocalizedData.WebsiteStateFailureError) -f ${Name} ;
                        $errorMessage += $_.Exception.Message
                        $exception = New-Object System.InvalidOperationException $errorMessage ;
                        $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

                        $PSCmdlet.ThrowTerminatingError($errorRecord)
                    }
                }

                Write-Verbose("State for website $Name has been updated to $State");

            }

            if($UpdateNotRequired)
            {
                Write-Verbose("Website $Name already exists and properties do not need to be udpated.");
            }
            

        }
        else #Website doesn't exist so create new one
        {
            try
            {
                $Website = New-Website @psboundparameters
                $Result = Stop-Website $Website.name -ErrorAction Stop
            
                #Clear default bindings if new bindings defined and are different
                if($BindingInfo -ne $null)
                {
                    if(ValidateWebsiteBindings -Name $Name -BindingInfo $BindingInfo)
                    {
                        UpdateBindings -Name $Name -BindingInfo $BindingInfo
                    }
                }

        #Add Default pages for new created website  
            if($DefaultPage -ne $null)
                {
                UpdateDefaultPages -Name $Name -DefaultPage $DefaultPage  
        }

                Write-Verbose("successfully created website $Name")
                
                #Start site if required
                if($State -eq "Started")
                {
                    #Wait 1 sec for bindings to take effect
                    #I have found that starting the website results in an error if it happens to quickly
                    Start-Sleep -Seconds 1
                    Start-Website -Name $Name -ErrorAction Stop
                }

                Write-Verbose("successfully started website $Name")
      
            }
            catch
           {
                $errorId = "WebsiteCreationFailure"; 
                $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidOperation;
                $errorMessage = $($LocalizedData.WebsiteCreationFailureError) -f ${Name} ;
                $errorMessage += $_.Exception.Message
                $exception = New-Object System.InvalidOperationException $errorMessage ;
                $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null
                $PSCmdlet.ThrowTerminatingError($errorRecord);        
            }
        }    
    }
    else #Ensure is set to "Absent" so remove website 
    { 
        try
        {
            $website = Get-WebSiteByName $Name
            if($website -ne $null)
            {
                Remove-Website -name $Name
        
                Write-Verbose("Successfully removed Website $Name.")
            }
            else
            {
                Write-Verbose("Website $Name does not exist.")
            }
        }
        catch
        {
            $errorId = "WebsiteRemovalFailure"; 
            $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidOperation;
            $errorMessage = $($LocalizedData.WebsiteRemovalFailureError) -f ${Name} ;
            $errorMessage += $_.Exception.Message
            $exception = New-Object System.InvalidOperationException $errorMessage ;
            $errorRecord = New-Object System.Management.Automation.ErrorRecord $exception, $errorId, $errorCategory, $null

            $PSCmdlet.ThrowTerminatingError($errorRecord);
        }
        
    }
}


# The Test-TargetResource cmdlet is used to validate if the role or feature is in a state as expected in the instance document.
function Test-TargetResource 
{
    [OutputType([System.Boolean])]
    param 
    (       
        [ValidateSet("Present", "Absent")]
        [string]$Ensure = "Present",

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PhysicalPath,

        [ValidateSet("Started", "Stopped")]
        [string]$State = "Started",

        [string]$ApplicationPool,

        [Microsoft.Management.Infrastructure.CimInstance[]]$BindingInfo,

    [string[]]$DefaultPage
    )
 
    $DesiredConfigurationMatch = $true;

    # Check if WebAdministration module is present for IIS cmdlets
    if(!(Get-Module -ListAvailable -Name WebAdministration))
    {
        Throw "Please ensure that WebAdministration module is installed."
    }

    $count = Test-WebSiteByName $Name
    $Stop = $true

    Do
    {
        #Check Ensure
        if(($Ensure -eq "Present" -and $count -eq 0) -or ($Ensure -eq "Absent" -and $count -ne 0))
        {
            $DesiredConfigurationMatch = $false
            Write-Verbose("The Ensure state for website $Name does not match the desired state.");
            break
        }

        # Only check properties if $website exists
        if ($count -ne 0)
        {
            $website = Get-WebSiteByName $Name

            #Check Physical Path property
            if(ValidateWebsitePath -Name $Name -PhysicalPath $PhysicalPath)
            {
                $DesiredConfigurationMatch = $false
                Write-Verbose("Physical Path of Website $Name does not match the desired state.");
                break
            }

            #Check State
            if($website.state -ne $State -and $State -ne $null)
            {
                $DesiredConfigurationMatch = $false
                Write-Verbose("The state of Website $Name does not match the desired state.");
                break
            }

            #Check Application Pool property 
            if(($ApplicationPool -ne "") -and ($website.applicationPool -ne $ApplicationPool))
            {
                $DesiredConfigurationMatch = $false
                Write-Verbose("Application Pool for Website $Name does not match the desired state.");
                break
            }

            #Check Binding properties
            if($BindingInfo -ne $null)
            {
                if(ValidateWebsiteBindings -Name $Name -BindingInfo $BindingInfo)
                {
                    $DesiredConfigurationMatch = $false
                    Write-Verbose("Bindings for website $Name do not mach the desired state.");
                    break
                }

            }
        }

        #Check Default Pages 
        if($DefaultPage -ne $null)
        {
            $allDefaultPage = @(Get-WebConfiguration //defaultDocument/files/*  -PSPath (Join-Path "IIS:\sites\" $Name) |%{Write-Output $_.value})

            $allDefaultPagesPresent = $true

                foreach($page in $DefaultPage )
                {
                    if(-not ($allDefaultPage  -icontains $page))
                    {   
                        $DesiredConfigurationMatch = $false
                        Write-Verbose("Default Page for website $Name do not mach the desired state.");
                        $allDefaultPagesPresent = $false  
                        break
                    }
                }
        
            if($allDefaultPagesPresent -eq $false)
            {
                # This is to break out from Test 
                break 
            }
        }


        $Stop = $false
    }
    While($Stop)   

    $DesiredConfigurationMatch;
}
