###########################################################
# SharePoint Farm Poster Script Generator
# Author : henri d'Orgeval
# This source is licenced under Microsoft Public License (Ms-PL)
# All other rights reserved
# This script is provided "as-is"
# Check for any update @ http://spposter.codeplex.com

$ScriptVersion="0.9.68"

###########################################################
#Core Functions
function GetArrayItemCount($array)
{
    if ( $array -eq $null)
    {
        return 0;
    }
    
    try {
        return $array.Count
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
    }
    
    return 0
}

function WriteProgress 
{
    param([String]$status, [int32]$percent)
    $progressActivity = "Generating Farm Poster v$ScriptVersion"
    Write-Progress -activity $progressActivity -status $status -percentComplete $percent
}

function WaitFor2Seconds()
{
    [System.Threading.Thread]::Sleep(2000)
}

function NewPropertySet()
{
    $args = @{} #Hash table
    return $args
}

function SafeToString($obj)
{
    if ($obj -eq $null)
    {
        return $null
    }
    $result = ""
    try {
        $result = $obj.ToString([System.Globalization.CultureInfo]::CurrentUICulture)
    }
    catch [Exception] {
         #$msg = $_.Exception.Message
         $result = $obj.ToString()
    }
    
    return $result
}


###########################################################
# Begin activity
WriteProgress "Loading SharePoint Snapin (may take several minutes) ..." 5
Add-PSSnapin Microsoft.SharePoint.PowerShell
WaitFor2Seconds
# End activity

###########################################################
# Begin activity
WriteProgress "Loading Core SharePoint functions..." 7

function GetLocalFarm()
{
    return [Microsoft.SharePoint.Administration.SPFarm]::Local
}

function GetSPWebServiceInstance()
{
    $farm = GetLocalFarm
    $instance = [Microsoft.SharePoint.Administration.SPWebService] $null
    foreach ($svc in  $farm.Services)
    {
        if ( ! $svc.GetType().equals([Microsoft.SharePoint.Administration.SPWebService]))
        {
            continue
        }
        $instance = [Microsoft.SharePoint.Administration.SPWebService] $svc
        break
    }
    return $instance
}

function GetAllWebApplications()
{
    $webApplications = Get-SPWebApplication -IncludeCentralAdministration
    
    $WebApplicationsProperty = @() # array of objects
    foreach ($webApplication in $webApplications)
    {
        $webApplicationProperties = @{} #Hash table
        $webApplicationProperties.Name = $webApplication.DisplayName
        
        #
        $webApplicationSettings = New-Object PSObject -Property $webApplicationProperties
        $WebApplicationsProperty +=  $webApplicationSettings 
    }
    return $WebApplicationsProperty
}


function GetAssociatedServiceApplicationClonedInstance([Microsoft.SharePoint.Administration.SPServiceApplicationProxy] $serviceApplicationProxy)
{
    $serviceApplicationProperties = @{} #Hash table
    $mustStopIteration = $false
    $farm = GetLocalFarm
    foreach ($svc in  $farm.Services)
    {
        foreach ($serviceApplication in $svc.Applications)
        {
            try {
                if ($ServiceApplication.IsConnected($serviceApplicationProxy))
                {
                    # Get Description of this Service Application
                    $serviceApplicationProperties.Type = $serviceApplication.TypeName
                    
                    # Get State of this proxy
                    $serviceApplicationProperties.Status = $serviceApplication.Status.ToString()
                    
                    # Get Admin Url in Central Admin to configure this Service Application
                    $serviceApplicationProperties.ManageServiceApplicationUrl = $null
                    if ( $serviceApplication.ManageLink -ne $null)
                    {
                        $serviceApplicationProperties.ManageServiceApplicationUrl = $centralAdminUrl + ($serviceApplication.ManageLink.Url)
                    }
                    $mustStopIteration = $true
                    break
                }
            }
            catch [Exception] {
                 #$msg = $_.Exception.Message
            }
        } # end foreach ($serviceApplication in $svc.Applications)
        
        
        if ($mustStopIteration -eq $true) { break }
    } # end foreach ($svc in  $SPFarm.Services)
    
    #
    $ClonedServiceApplication = New-Object PSObject -Property $serviceApplicationProperties
    return $ClonedServiceApplication
}


function GetProxyClonedInstance([Microsoft.SharePoint.Administration.SPServiceApplicationProxy]$proxy)
{
    $proxyProperties = @{} #Hash table
    $proxyProperties.Name = $proxy.DisplayName
    $proxyProperties.Id = $proxy.Id
    $proxyProperties.DotNetType = $proxy.GetType().ToString()
    
    # Get Admin Url in Central Admin to configure this proxy
    $proxyProperties.ManageProxyUrl = $null
    if ( $proxy.PropertiesLink -ne $null)
    {
        $proxyProperties.ManageProxyUrl = $centralAdminUrl + ($proxy.PropertiesLink.Url)
    }
    
    # Get Description of this proxy
    $proxyProperties.Type = $proxy.TypeName
    
    # Get State of this proxy
    $proxyProperties.Status = $proxy.Status.ToString()
    
    # Get Parent Service Application
    $proxyProperties.ServiceApplication = GetAssociatedServiceApplicationClonedInstance($proxy)
    
    #
    $ClonedProxy = New-Object PSObject -Property $proxyProperties
    return  $ClonedProxy
}

function GetQuotaTemplateClonedInstance([Microsoft.SharePoint.Administration.SPQuotaTemplate]$quotaTemplate)
{
    $quotaTemplateProperties = @{} #Hash table
    $quotaTemplateProperties.Name = $quotaTemplate.Name
    
    $maxDiskUsage = $quotaTemplate.StorageMaximumLevel
    $maxDiskUsageInGB = ([Double]$maxDiskUsage)/[Double](1024*1024)
    #
    $storageWarningLevel = $quotaTemplate.StorageWarningLevel
    $storageWarningLevelInGB = ([Double]$storageWarningLevel)/[Double](1024*1024)
    #
    $quotaTemplateProperties.MaxDiskSpaceForSiteCollection = [string]::Format("{0:F0} Mb (Send warning e-mail @ {1:F0} Mb)",$maxDiskUsageInGB,$storageWarningLevelInGB) 
    
    $maxResourceUsageForSiteCollection = $quotaTemplate.UserCodeMaximumLevel
    $userCodeWarningLevel = $quotaTemplate.UserCodeWarningLevel
    $quotaTemplateProperties.MaxResourceUsageForSiteCollection = [string]::Format("{0} points (Send warning e-mail @ {1} points)",$maxResourceUsageForSiteCollection,$userCodeWarningLevel) 
    #
    $ClonedQuotaTemplate = New-Object PSObject -Property $quotaTemplateProperties
    return $ClonedQuotaTemplate

}

function GetFeatureDefinitionClonedInstance([Microsoft.SharePoint.Administration.SPFeatureDefinition]$f)
{
    $featureProperties = @{} #Hash table
    $internalName = $f.DisplayName
    $featureProperties.Name = $f.DisplayName
    $featureProperties.Title = $f.GetTitle([System.Globalization.CultureInfo]::CurrentUICulture)
    $featureProperties.Description = $f.GetDescription([System.Globalization.CultureInfo]::CurrentUICulture)
    $featureProperties.GUID = $f.Id
    $featureProperties.Hidden = $f.Hidden
    $featureProperties.Version = "Not Set"
    $featureProperties.VersionOnDisk = "Not Set"
    
    if ( $f.Version -ne $null)
    {
        $featureProperties.VersionOnDisk = SafeToString($f.Version)
        $featureProperties.Version = SafeToString($f.Version)
    }
    
    # check if Feature exists on disk
    #$setupPath = $f.RootDirectory + "\feature.xml"
    #$featureProperties.ExistsOnDisk = $false
    #if ([System.IO.Directory]::Exists($f.RootDirectory))
    #{
    #    if ([System.IO.File]::Exists($setupPath))
    #    {
    #        $featureProperties.ExistsOnDisk = $true
    #    }
    #}
    
    $featureProperties.RootDirectory = $f.RootDirectory.Replace("$sharepointRoot" , "{SharePointRoot}\")
    $featureProperties.RequireResources = $f.RequireResources
    $featureProperties.DefaultResourceFile = $f.DefaultResourceFile
    $featureProperties.IsActivated = $false
    $spWebService = GetSPWebServiceInstance
    if ($spWebService.Features[$f.Id] -ne $null )
    {
        $featureProperties.IsActivated = $true
        # Get the Version of Activated Feature
        $featureProperties.Version = SafeToString($spWebService.Features[$f.Id].Version)
    }
    
    $ClonedFeatureDefinition = New-Object PSObject -Property $featureProperties
    return $ClonedFeatureDefinition
}

function GetDeveloperDashboardSettingsClonedInstance()
{
    $spWebService = GetSPWebServiceInstance
    $developerDashboardSettings = $spWebService.DeveloperDashboardSettings

    $developerDashboardProperties = @{} #Hash table
    $developerDashboardProperties.ShowLinkToDisplayFullVerboseTrace = $developerDashboardSettings.TraceEnabled
    $developerDashboardProperties.RequiredPermissions = $developerDashboardSettings.RequiredPermissions.ToString()
    $developerDashboardProperties.DisplayLevel = $developerDashboardSettings.DisplayLevel.ToString()

    $ClonedDeveloperDashboardSettings = New-Object PSObject -Property $developerDashboardProperties
    return $ClonedDeveloperDashboardSettings
}

function GetDiagnosticsServiceClonedInstance()
{
    $spLogSettings = [Microsoft.SharePoint.Administration.SPDiagnosticsService]::Local
    $spLogSettingsProperties = @{} #Hash table
    $spLogSettingsProperties.CEIPEnabled = $spLogSettings.CEIPEnabled
    $spLogSettingsProperties.DaysToKeepLogs = $spLogSettings.DaysToKeepLogs
    $spLogSettingsProperties.MinutesOfTracingPerLogFile = $spLogSettings.LogCutInterval
    $spLogSettingsProperties.MaxDiskUsageInGB = $spLogSettings.LogDiskSpaceUsageGB
    $spLogSettingsProperties.MaxNumberOfLogFilesToKeep = $spLogSettings.LogsToKeep
    $spLogSettingsProperties.RestrictDiskSpaceUsage = $spLogSettings.LogMaxDiskSpaceUsageEnabled
    $spLogSettingsProperties.Location = $spLogSettings.LogLocation

    $ClonedLogSettings = New-Object PSObject -Property $spLogSettingsProperties
    return $ClonedLogSettings
}


function GetDeployedServersForFarmSolution([Microsoft.SharePoint.Administration.SPSolution] $solution)
{
    $servers = $solution.DeployedServers
    $result = ""
    $count = 0
    foreach ($server in $servers)
    {
        if ( $count -gt 0 )
        {
            $result += ", "
        }
        $serverName = $server.Name
        $result += $serverName
        $count += 1
    }
    return $result
}

function GetSolutionClonedInstance([Microsoft.SharePoint.Administration.SPSolution] $solution)
{
    $solutionProperties = @{} #Hash table
    #
    $solutionProperties.Name = $solution.Name
    $solutionProperties.Id = $solution.SolutionId.ToString()
    $solutionProperties.Deployed = $solution.Deployed
    $solutionProperties.ContainsCodeDeployedToGAC = $solution.ContainsGlobalAssembly
    #
    $deployedServers = GetDeployedServersForFarmSolution($solution)
    $solutionProperties.DeployedServers = $deployedServers

    $ClonedSolution = New-Object PSObject -Property $solutionProperties
    return $ClonedSolution
}

function GetAllServicesRunByThisAccount($accountName)
{

    if ( $accountName -eq $null)
    {
        return $null
    }
    
    $collection = @() # array of objects
    $farm = GetLocalFarm
    $farmServices = $farm.Services
    foreach ($svc in  $farmServices)
    {
        $spWindowsService = $svc -as [Microsoft.SharePoint.Administration.SPWindowsService]
        if ( $spWindowsService -ne $null)
        {
            # current SharePoint Service is a Windows Service
            $processIdentity = $spWindowsService.ProcessIdentity
            if ( $processIdentity -eq $null)
            {
                continue
            }

            $userName = $processIdentity.UserName
            
            if ( $userName -eq $null)
            {
                #might be a bug in Microsoft implementation : sometimes processIdentity returns a SPProcessIdentity object, sometimes it returns a string
                $userName = $processIdentity -as [System.String]
                if ($userName -eq $null)
                {
                    continue
                }
            }
            
            if ($userName -eq $null)
            {
                continue
            }

            if ($userName.Equals($accountName))
            {
                $serviceProperties = @{} #Hash table
                $serviceProperties.Description = $spWindowsService.TypeName
                $serviceProperties.Name = $spWindowsService.Name
                $serviceProperties.Type = "Windows Service"
                $ClonedService = New-Object PSObject -Property $serviceProperties
                $collection += $ClonedService
            }
            
            if ($processIdentity.CurrentIdentityType -eq $null)
            {
                continue
            }
            
            if ($processIdentity.CurrentIdentityType.Equals([Microsoft.SharePoint.Administration.IdentityType]::LocalService))
            {
                if ($accountName.Equals("LocalService") )
                {
                    $serviceProperties = @{} #Hash table
                    $serviceProperties.Description = $spWindowsService.TypeName
                    $serviceProperties.Name = $spWindowsService.Name
                    $serviceProperties.Type = "Windows Service"
                    $ClonedService = New-Object PSObject -Property $serviceProperties
                    $collection += $ClonedService
                }
            }
            if ($processIdentity.CurrentIdentityType.Equals([Microsoft.SharePoint.Administration.IdentityType]::LocalSystem))
            {
                if ($accountName.Equals("LocalSystem") )
                {
                    $serviceProperties = @{} #Hash table
                    $serviceProperties.Description = $spWindowsService.TypeName
                    $serviceProperties.Name = $spWindowsService.Name
                    $serviceProperties.Type = "Windows Service"
                    $ClonedService = New-Object PSObject -Property $serviceProperties
                    $collection += $ClonedService
                }
            }
            if ($processIdentity.CurrentIdentityType.Equals([Microsoft.SharePoint.Administration.IdentityType]::NetworkService))
            {
                if ($accountName.Equals("NetworkService") )
                {
                    $serviceProperties = @{} #Hash table
                    $serviceProperties.Description = $spWindowsService.TypeName
                    $serviceProperties.Name = $spWindowsService.Name
                    $serviceProperties.Type = "Windows Service"
                    $ClonedService = New-Object PSObject -Property $serviceProperties
                    $collection += $ClonedService
                }
            }
            continue
        } # end if ( $spWindowsService -ne $null)
        
        $spWebService = $svc -as [Microsoft.SharePoint.Administration.SPWebService]
        if ( $spWebService -ne $null)
        {
            # current SharePoint Service is a SPWebService
            # get all WebApp 
            $spWebApplications = $spWebService.WebApplications
            foreach ($spWebApplication in $spWebApplications)
            {
                $userName = $spWebApplication.ApplicationPool.UserName
                if ( $userName -eq $null)
                {
                    continue
                }
                if ($userName.Equals($accountName))
                {
                    $serviceProperties = @{} #Hash table
                    $serviceProperties.Description = $spWebApplication.DisplayName
                    $serviceProperties.Name = $spWebApplication.ApplicationPool.DisplayName
                    $serviceProperties.Type = "Web Application"
                    $ClonedService = New-Object PSObject -Property $serviceProperties
                    $collection += $ClonedService
                }
                $processIdentity = $spWebApplication.ApplicationPool.ProcessIdentity
                if ( $processIdentity -eq $null)
                {
                    continue
                }
                if ($processIdentity.CurrentIdentityType.Equals([Microsoft.SharePoint.Administration.IdentityType]::LocalService))
                {
                    if ($accountName.Equals("LocalService") )
                    {
                        $serviceProperties = @{} #Hash table
                        $serviceProperties.Description = $spWebApplication.DisplayName
                        $serviceProperties.Name = $spWebApplication.ApplicationPool.DisplayName
                        $serviceProperties.Type = "Web Application"
                        $ClonedService = New-Object PSObject -Property $serviceProperties
                        $collection += $ClonedService
                    }
                }
                if ($processIdentity.CurrentIdentityType.Equals([Microsoft.SharePoint.Administration.IdentityType]::LocalSystem))
                {
                    if ($accountName.Equals("LocalSystem") )
                    {
                        $serviceProperties = @{} #Hash table
                        $serviceProperties.Description = $spWebApplication.DisplayName
                        $serviceProperties.Name = $spWebApplication.ApplicationPool.DisplayName
                        $serviceProperties.Type = "Web Application"
                        $ClonedService = New-Object PSObject -Property $serviceProperties
                        $collection += $ClonedService
                    }
                }
                if ($processIdentity.CurrentIdentityType.Equals([Microsoft.SharePoint.Administration.IdentityType]::NetworkService))
                {
                    if ($accountName.Equals("NetworkService") )
                    {
                        $serviceProperties = @{} #Hash table
                        $serviceProperties.Description = $spWebApplication.DisplayName
                        $serviceProperties.Name = $spWebApplication.ApplicationPool.DisplayName
                        $serviceProperties.Type = "Web Application"
                        $ClonedService = New-Object PSObject -Property $serviceProperties
                        $collection += $ClonedService
                    }
                }
                
            } # end foreach ($spWebApplication in $spWebApplications)
            
            
            continue
        } # end if ( $spWebService -ne $null)
        
        
    } # end foreach ($svc in  $SPFarm.Services)
    
    return $collection
}

function GetManagedAccountClonedInstance($account)
{
    $managedAccountProperties = @{} #Hash table
    
    $managedAccount = $account -as [Microsoft.SharePoint.Administration.SPManagedAccount]
    if ($managedAccount -ne $null)
    {
        $managedAccountProperties.UserName = $managedAccount.UserName
        #$managedAccountProperties.AccountType = $account.TypeName
        $managedAccountProperties.EnableSharePointToAutomaticallyGenerateAndUpdatePassword = $managedAccount.AutomaticChange
        $managedAccountProperties.PasswordLastChange = SafeToString($managedAccount.PasswordLastChanged)
        $managedAccountProperties.PasswordExpirationDate = SafeToString($managedAccount.PasswordExpiration)
        $managedAccountProperties.PasswordChangeSchedule = SafeToString($managedAccount.ChangeSchedule)
    }
    
    if ($managedAccount -eq $null)
    {
        $managedAccountProperties.UserName = $account
        $managedAccountProperties.EnableSharePointToAutomaticallyGenerateAndUpdatePassword = $false
    }
    
    # Get all services that are run by this account
    $managedAccountProperties.Services = GetAllServicesRunByThisAccount($managedAccountProperties.UserName)
    
    $ClonedManagedAccount = New-Object PSObject -Property $managedAccountProperties
    $properties.ManagedAccounts +=  $ClonedManagedAccount 
} # end function AddAccount($account)


function GetWebApplicationGlobalSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $GlobalSettingsProperties = @{} #Hash table
    $GlobalSettingsProperties.DefaultQuotaTemplate = $webApplication.DefaultQuotaTemplate
    $GlobalSettingsProperties.ServiceApplicationProxyGroup = $webApplication.ServiceApplicationProxyGroup.FriendlyName
    $GlobalSettingsProperties.BrowserCEIPEnabled = $webApplication.BrowserCEIPEnabled
    $GlobalSettingsProperties.SizeOfTheLargestFileThatCanBeUploaded = $webApplication.MaximumFileSize.ToString() + " MB"
    $GlobalSettingsProperties.SelfServiceSiteCreationEnabled = $webApplication.SelfServiceSiteCreationEnabled
    $GlobalSettingsProperties.RecycleBinRetentionPeriod = $webApplication.RecycleBinRetentionPeriod.ToString() + " days"
    #
    $ClonedGlobalSettings = New-Object PSObject -Property $GlobalSettingsProperties
    return  $ClonedGlobalSettings
}


function GetInlineDownloadedMimeTypesSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $InlineDownloadedMimeTypesProperties = @{} #Hash table
    $InlineDownloadedMimeTypesProperties.Description = "list of MIME content-types for this application which are allowed to be viewed without first saving to the user's localdrive"
    $InlineDownloadedMimeTypesProperties.MimeTypes = "Not Set"
    if ($webapplication.AllowedInlineDownloadedMimeTypes -eq $null)
    {
        $ClonedInlineDownloadedMimeTypesSettings = New-Object PSObject -Property $InlineDownloadedMimeTypesProperties
        return  $ClonedInlineDownloadedMimeTypesSettings 
    }
    
    $InlineDownloadedMimeTypesProperties.MimeTypes = ""
    $count = 0
    foreach ($mime in $webApplication.AllowedInlineDownloadedMimeTypes)
    {
        if ($count -gt 0)
        {
            $InlineDownloadedMimeTypesProperties.MimeTypes += ", "
        }
        $InlineDownloadedMimeTypesProperties.MimeTypes += $mime.ToString()
        $count += 1
    }
    
    $ClonedInlineDownloadedMimeTypesSettings = New-Object PSObject -Property $InlineDownloadedMimeTypesProperties
    return  $ClonedInlineDownloadedMimeTypesSettings 

}

function GetBlockedFileExtensionsSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $BlockedFileExtensionsProperties = @{} #Hash table
    $BlockedFileExtensionsProperties.Description = "List of file extensions that cannot be uploaded or downloaded from sites in the Web application"
    $BlockedFileExtensionsProperties.BlockedFileExtensions = "Not Set"
    if ($webapplication.BlockedFileExtensions -eq $null)
    {
        $ClonedBlockedFileExtensionsSettings = New-Object PSObject -Property $BlockedFileExtensionsProperties
        return  $ClonedBlockedFileExtensionsSettings 
    }
    
    $BlockedFileExtensionsProperties.BlockedFileExtensions = ""
    $count = 0
    foreach ($extension in $webApplication.BlockedFileExtensions)
    {
        if ($count -gt 0)
        {
            $BlockedFileExtensionsProperties.BlockedFileExtensions += ", "
        }
        $BlockedFileExtensionsProperties.BlockedFileExtensions += "."
        $BlockedFileExtensionsProperties.BlockedFileExtensions += $extension.ToString()
        $count += 1
    }
    
    $ClonedBlockedFileExtensionsSettings = New-Object PSObject -Property $BlockedFileExtensionsProperties
    return  $ClonedBlockedFileExtensionsSettings  

}

function GetManagedPathSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $ManagedPathProperties = @() # array of objects
    $managedPaths = $webApplication.Prefixes
    foreach ($managedPath in $managedPaths)
    {
        $ManagedPathProperty = @{} #Hash table
        $ManagedPathProperty.Path = "/" + $managedPath.Name
        $ManagedPathProperty.PathType = SafeToString($managedPath.PrefixType)
        $ManagedPathProperty.Description = "Only one Site Collection can be created under this path" 
        if ($ManagedPathProperty.PathType.Equals("WildcardInclusion"))
        {
            $ManagedPathProperty.Description = "As many Site Collection as needed can be created under this path" 
        }
        $ClonedManagedPath = New-Object PSObject -Property $ManagedPathProperty
        $ManagedPathProperties +=  $ClonedManagedPath 
    }
    return $ManagedPathProperties
}

function GetLargeListThrottlingSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $LargeListThrotlingProperties = @{} #Hash table
    $LargeListThrotlingProperties.MaxItemsPerThrottledOperation = $webApplication.MaxItemsPerThrottledOperation
    $LargeListThrotlingProperties.MaxItemsPerThrottledOperationForAdministratorOrAuditor = $webApplication.MaxItemsPerThrottledOperationOverride
    $LargeListThrotlingProperties.MaxItemsWithUniquePermissions = $webApplication.MaxUniquePermScopesPerList
    $LargeListThrotlingProperties.AllowCustomCodeToOverrideListThrottlingSettings = $webApplication.AllowOMCodeOverrideThrottleSettings
    $LargeListThrotlingProperties.EnableUnthrottledDailyTimeWindow = $webApplication.UnthrottledPrivilegedOperationWindowEnabled
    if ( $webApplication.UnthrottledPrivilegedOperationWindowEnabled -eq $true )
    {
        $LargeListThrotlingProperties.StartTimeForTheUnthrottledOperations = [string]::Format("{0}h {1}minute",$webApplication.DailyStartUnthrottledPrivilegedOperationsHour,$webApplication.DailyStartUnthrottledPrivilegedOperationsMinute)
        $LargeListThrotlingProperties.DurationInHoursForTheUnthrottledOperations = $webApplication.DailyUnthrottledPrivilegedOperationsDuration
    }
    
    $ClonedLargeListThrottlingSettings = New-Object PSObject -Property $LargeListThrotlingProperties
    return $ClonedLargeListThrottlingSettings
}


function GetIISSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $IISSettingsProperties = @() # array of objects
    foreach ($urlZone in $webApplication.IisSettings.Keys)
    {
        $iisSettings = $webApplication.IisSettings[$urlZone]
        $UrlZoneProperties = @{} #Hash table
        $UrlZoneProperties.Zone = $urlZone.ToString()
        $UrlZoneProperties.Port = $iisSettings.ServerBindings[0].Port
        $UrlZoneProperties.HostHeader = $iisSettings.ServerBindings[0].HostHeader
        $UrlZoneProperties.PublicationFolder = $iisSettings.Path.ToString()
        # Get the public URL
        foreach ( $alternateUrl in $webApplication.AlternateUrls)
        {
            if ( $alternateUrl.UrlZone -eq $urlZone )
            {
                $UrlZoneProperties.Url = $alternateUrl.IncomingUrl
                break
            }
        }
        
        # Get Authentication details
        $AuthenticationProperties = @{} #Hash table
        $AuthenticationProperties.AuthenticationMode = $iisSettings.AuthenticationMode.ToString()
        $AuthenticationProperties.AllowAnonymous = $iisSettings.AllowAnonymous
        $AuthenticationProperties.UseBasicAuthentication = $iisSettings.UseBasicAuthentication
        $AuthenticationProperties.UseClaimsAuthentication = $iisSettings.UseClaimsAuthentication
        $AuthenticationProperties.UseFormsClaimsAuthenticationProvider = $iisSettings.UseFormsClaimsAuthenticationProvider
        $AuthenticationProperties.UseClaimsAuthentication = $iisSettings.UseClaimsAuthentication
        $AuthenticationProperties.UseTrustedClaimsAuthenticationProvider = $iisSettings.UseTrustedClaimsAuthenticationProvider
        $AuthenticationProperties.UseWindowsClaimsAuthenticationProvider = $iisSettings.UseWindowsClaimsAuthenticationProvider
        $AuthenticationProperties.UseWindowsIntegratedAuthentication = $iisSettings.UseWindowsIntegratedAuthentication
        $AuthenticationProperties.DisableKerberos = $iisSettings.DisableKerberos
        $AuthenticationProperties.EnableClientIntegration = $iisSettings.EnableClientIntegration

        $ClonedAuthenticationSettings = New-Object PSObject -Property $AuthenticationProperties
        
        $UrlZoneProperties.AuthenticationSettings = $ClonedAuthenticationSettings
        
        $ClonedIisSettings = New-Object PSObject -Property $UrlZoneProperties
        $IISSettingsProperties +=  $ClonedIisSettings 
    }
    return $IISSettingsProperties
}

function GetSharePointDesignerSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $SharePointDesignerProperties = @{} #Hash table
    $SharePointDesignerProperties.AllowSharePointDesigner = $webApplication.AllowDesigner
    $SharePointDesignerProperties.AllowMasterPageEditing = $webApplication.AllowMasterPageEditing
    $SharePointDesignerProperties.AllowCustomizedSiteToBeRevertedToTheirBaseTemplate = $webApplication.AllowRevertFromTemplate
    $SharePointDesignerProperties.AllowUsersToSeeFileStructureOfWebSite = $webApplication.ShowURLStructure

    $ClonedSharePointDesignerSettings = New-Object PSObject -Property $SharePointDesignerProperties
    return  $ClonedSharePointDesignerSettings
}

function GetBasePermissionDescription([Microsoft.SharePoint.SPBasePermissions]$basePermission)
{
    $desc = ""
    $AddAndCustomizePages = [Microsoft.SharePoint.SPBasePermissions]::AddAndCustomizePages
    $AddDelPrivateWebParts  = [Microsoft.SharePoint.SPBasePermissions]::AddDelPrivateWebParts
    $AddListItems  = [Microsoft.SharePoint.SPBasePermissions]::AddListItems
    $ApplyStyleSheets  = [Microsoft.SharePoint.SPBasePermissions]::ApplyStyleSheets
    $ApplyThemeAndBorder  = [Microsoft.SharePoint.SPBasePermissions]::ApplyThemeAndBorder
    $ApproveItems  = [Microsoft.SharePoint.SPBasePermissions]::ApproveItems
    $BrowseDirectories  = [Microsoft.SharePoint.SPBasePermissions]::BrowseDirectories
    $BrowseUserInfo  = [Microsoft.SharePoint.SPBasePermissions]::BrowseUserInfo
    $BrowseDirectories  = [Microsoft.SharePoint.SPBasePermissions]::BrowseDirectories
    $BrowseUserInfo  = [Microsoft.SharePoint.SPBasePermissions]::BrowseUserInfo
    $CancelCheckout  = [Microsoft.SharePoint.SPBasePermissions]::CancelCheckout
    $CreateAlerts  = [Microsoft.SharePoint.SPBasePermissions]::CreateAlerts
    $CreateGroups  = [Microsoft.SharePoint.SPBasePermissions]::CreateGroups
    $CreateSSCSite  = [Microsoft.SharePoint.SPBasePermissions]::CreateSSCSite
    $DeleteListItems  = [Microsoft.SharePoint.SPBasePermissions]::DeleteListItems
    $DeleteVersions  = [Microsoft.SharePoint.SPBasePermissions]::DeleteVersions
    $EditListItems  = [Microsoft.SharePoint.SPBasePermissions]::EditListItems
    $EditMyUserInfo  = [Microsoft.SharePoint.SPBasePermissions]::EditMyUserInfo
    $EnumeratePermissions  = [Microsoft.SharePoint.SPBasePermissions]::EnumeratePermissions
    $ManageAlerts  = [Microsoft.SharePoint.SPBasePermissions]::ManageAlerts
    $ManageLists  = [Microsoft.SharePoint.SPBasePermissions]::ManageLists
    $ManagePermissions  = [Microsoft.SharePoint.SPBasePermissions]::ManagePermissions
    $ManagePersonalViews  = [Microsoft.SharePoint.SPBasePermissions]::ManagePersonalViews
    $ManageSubwebs  = [Microsoft.SharePoint.SPBasePermissions]::ManageSubwebs
    $ManageWeb  = [Microsoft.SharePoint.SPBasePermissions]::ManageWeb
    $Open  = [Microsoft.SharePoint.SPBasePermissions]::Open
    $OpenItems  = [Microsoft.SharePoint.SPBasePermissions]::OpenItems
    $UpdatePersonalWebParts  = [Microsoft.SharePoint.SPBasePermissions]::UpdatePersonalWebParts
    $UseClientIntegration  = [Microsoft.SharePoint.SPBasePermissions]::UseClientIntegration
    $UseRemoteAPIs  = [Microsoft.SharePoint.SPBasePermissions]::UseRemoteAPIs
    $ViewFormPages  = [Microsoft.SharePoint.SPBasePermissions]::ViewFormPages
    $ViewListItems  = [Microsoft.SharePoint.SPBasePermissions]::ViewListItems
    $ViewPages  = [Microsoft.SharePoint.SPBasePermissions]::ViewPages
    $ViewUsageData  = [Microsoft.SharePoint.SPBasePermissions]::ViewUsageData
    $ViewVersions  = [Microsoft.SharePoint.SPBasePermissions]::ViewVersions
    
    if ($basePermission -eq $AddAndCustomizePages)
    {
        return "Add, change, or delete HTML pages or Web Part Pages, and edit the Web site using a SharePoint Foundation–compatible editor."
    }
    
    if ($basePermission -eq $AddDelPrivateWebParts)
    {
        return "Add or remove personal Web Parts on a Web Part Page."
    }
    
    if ($basePermission -eq $AddListItems)
    {
        return "Add items to lists, add documents to document libraries, and add Web discussion comments."
    }
    
    if ($basePermission -eq $ApplyStyleSheets)
    {
        return "Apply a style sheet (.css file) to the Web site."
    }
    
    if ($basePermission -eq $ApplyThemeAndBorder)
    {
        return "Apply a theme or borders to the entire Web site."
    }
    
    if ($basePermission -eq $ApproveItems )
    {
        return "Approve a minor version of a list item or document."
    }
    
    if ($basePermission -eq $BrowseDirectories )
    {
        return "Enumerate files and folders in a Web site using Microsoft Office SharePoint Designer 2010 and WebDAV interfaces."
    }
    
    if ($basePermission -eq $BrowseUserInfo  )
    {
        return "View information about users of the Web site."
    }
    
    if ($basePermission -eq $CancelCheckout   )
    {
        return "Discard or check in a document which is checked out to another user."
    }
    
    if ($basePermission -eq $CreateAlerts    )
    {
        return "Create e-mail alerts."
    }
    
    if ($basePermission -eq $CreateGroups )
    {
        return "Create a group of users that can be used anywhere within the site collection."
    }
    
    if ($basePermission -eq $CreateSSCSite  )
    {
        return "Create a Web site using Self-Service Site Creation."
    }
    
    if ($basePermission -eq $DeleteListItems   )
    {
        return "Delete items from a list, documents from a document library, and Web discussion comments in documents."
    }
    
    if ($basePermission -eq $DeleteVersions   )
    {
        return "Delete past versions of a list item or document."
    }
    
    if ($basePermission -eq $EditListItems    )
    {
        return "Edit items in lists, edit documents in document libraries, edit Web discussion comments in documents, and customize Web Part Pages in document libraries."
    }
    
    if ($basePermission -eq $EditMyUserInfo    )
    {
        return "Allows a user to change his or her user information, such as adding a picture."
    }
    
    if ($basePermission -eq $EnumeratePermissions    )
    {
        return "Enumerate permissions on the Web site, list, folder, document, or list item."
    }
    
    if ($basePermission -eq $ManageAlerts    )
    {
        return "Manage alerts for all users of the Web site."
    }
    
    if ($basePermission -eq $ManageLists )
    {
        return "Create and delete lists, add or remove columns in a list, and add or remove public views of a list."
    }
    
    if ($basePermission -eq $ManagePermissions     )
    {
        return "Create and change permission levels on the Web site and assign permissions to users and groups."
    }
    
    if ($basePermission -eq $ManagePersonalViews  )
    {
        return "Create, change, and delete personal views of lists."
    }
    
    if ($basePermission -eq $ManageSubwebs  )
    {
        return "Create subsites such as team sites, Meeting Workspace sites, and Document Workspace sites."
    }
    
    if ($basePermission -eq $ManageWeb   )
    {
        return "Grant the ability to perform all administration tasks for the Web site as well as manage content."
    }
    
    if ($basePermission -eq $Open  )
    {
        return "Allow users to open a Web site, list, or folder to access items inside that container."
    }
    
    if ($basePermission -eq $OpenItems   )
    {
        return "View the source of documents with server-side file handlers."
    }
    
    if ($basePermission -eq $UpdatePersonalWebParts    )
    {
        return "Update Web Parts to display personalized information."
    }
    
    if ($basePermission -eq $UseClientIntegration     )
    {
        return "Use features that launch client applications; otherwise, users must work on documents locally and upload changes. "
    }
    
    if ($basePermission -eq $UseRemoteAPIs   )
    {
        return "Use SOAP, WebDAV, or Microsoft Office SharePoint Designer 2010 interfaces to access the Web site."
    }
    
    if ($basePermission -eq $ViewFormPages     )
    {
        return "View forms, views, and application pages, and enumerate lists."
    }
    
    if ($basePermission -eq $ViewListItems      )
    {
        return "View items in lists, documents in document libraries, and view Web discussion comments."
    }
    
    if ($basePermission -eq $ViewPages      )
    {
        return "View pages in a Web site."
    }
    
    if ($basePermission -eq $ViewUsageData      )
    {
        return "View reports on Web site usage."
    }
    
    if ($basePermission -eq $ViewVersions  )
    {
        return "View past versions of a list item or document."
    }
    
    return $desc
}

function GetWebPartScriptableSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $WebPartScriptableSettingsProperties = @{} #Hash table
    $WebPartScriptableSettingsProperties.AllowContributorsToEditScriptableParts = $webApplication.AllowContributorsToEditScriptableParts
    $WebPartScriptableSettingsProperties.Description = "A scriptable WebPart is any Web Part for which the safeControl entry in the web.config file contains the following attribute : SafeAgainstScript=""False"" "
    
    #
    $ClonedWebPartScriptableSettings = New-Object PSObject -Property $WebPartScriptableSettingsProperties
    return  $ClonedWebPartScriptableSettings
}

function GetWebPartSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $WebPartSettingsProperties = @{} #Hash table
    $WebPartSettingsProperties.AllowWebPartsToCommunicateWithEachOther = $webApplication.AllowPartToPartCommunication
    $WebPartSettingsProperties.AllowAccessToWebPartCatalog = $webApplication.AllowAccessToWebPartCatalog

    $WebPartSettingsProperties.ScriptableSettings = GetWebPartScriptableSettings($webApplication)
    #
    $ClonedWebPartSettings = New-Object PSObject -Property $WebPartSettingsProperties
    return  $ClonedWebPartSettings
}

function GetFullBasePermissions([Microsoft.SharePoint.SPBasePermissions]$basePermission)
{
    $BasePermissionsCollectionProperty = @() # array of objects
    $enumValues = [System.Enum]::GetValues([Microsoft.SharePoint.SPBasePermissions])
    foreach ($permissionValue in $enumValues)
    {
        $basePermissionProperties = @{} #Hash table
        $basePermissionProperties.Name = $permissionValue.ToString()
        $basePermissionProperties.Description = GetBasePermissionDescription($permissionValue)
        $emptymask = [Microsoft.SharePoint.SPBasePermissions]::EmptyMask
        $fullMask = [Microsoft.SharePoint.SPBasePermissions]::FullMask
        
        #do not record special value like EmptyMask
        if ($permissionValue -eq $emptymask)
        {
            continue
        }
        
        #do not record special value like FullMask
        if ($permissionValue -eq $fullMask)
        {
            continue
        }
        
        $basePermissionProperties.Enabled = $false 
        if ($basePermission -eq $emptymask)
        {
            $ClonedBasePermissionSettings = New-Object PSObject -Property $basePermissionProperties
            $BasePermissionsCollectionProperty +=  $ClonedBasePermissionSettings 
            continue
        }
        
        if ($basePermission -eq $fullMask)
        {
            $basePermissionProperties.Enabled = $true 
            $ClonedBasePermissionSettings = New-Object PSObject -Property $basePermissionProperties
            $BasePermissionsCollectionProperty +=  $ClonedBasePermissionSettings 
            continue
        }
        
        if ( ($permissionValue -band $basePermission) -eq $permissionValue)
        {
            $basePermissionProperties.Enabled = $true
        }
        $ClonedBasePermissionSettings = New-Object PSObject -Property $basePermissionProperties
        $BasePermissionsCollectionProperty +=  $ClonedBasePermissionSettings 
        
    } #foreach ($permissionValue in $enumValues)
    
    return $BasePermissionsCollectionProperty
}


function GetPermissionsSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $PermissionsSettingsProperties = @{} #Hash table
    $PermissionsSettingsProperties.RightsMask = GetFullBasePermissions($webApplication.RightsMask)
    #
    $ClonedPermissionsSettings = New-Object PSObject -Property $PermissionsSettingsProperties
    return  $ClonedPermissionsSettings
}

function GetDiskUsageForAllSiteCollections([Microsoft.SharePoint.Administration.SPContentDatabase]$contentDatabase)
{
    $size = [long] 0
    $sizeAsString = ""
    try {
        foreach ($siteCollection in $contentDatabase.Sites)
        {
            if ($siteCollection.Usage -ne $null)
            {
                $size += $siteCollection.Usage.Storage
            }
            $siteCollection.Dispose()
        }
        $sizeInGB = ([Double]$size)/[Double](1024*1024*1024)
        $sizeAsString = [string]::Format("{0:F2} Gb",$sizeInGB) 
    }
    catch [Exception] {
         $sizeAsString = $_.Exception.Message
    }
    
    return $sizeAsString
}


function GetContentDatabasesSettings([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $ContentDatabasesProperty = @() # array of objects
    $contentDatabases = $webApplication.ContentDatabases
    
    foreach ($db in $contentDatabases)
    {
        $contentDatabaseProperties = @{} #Hash table
        $contentDatabaseProperties.Name = $db.Name
        $contentDatabaseProperties.DatabaseServer = $db.Server
        $maxSiteCount = $db.MaximumSiteCount
        $warningSiteCount = $db.WarningSiteCount
        $currentSiteCount = $db.CurrentSiteCount
        
        $contentDatabaseProperties.CurrentNumberOfSiteCollection = [string]::Format("{0} (max = {1} / Send warning e-mail @ {2})",$currentSiteCount,$maxSiteCount,$warningSiteCount) 
        $contentDatabaseProperties.DiskUsageForAllSiteCollections = GetDiskUsageForAllSiteCollections($db)
        $contentDatabaseProperties.ConnectionString = $db.DatabaseConnectionString

        #get all Site collections in Content Database
        if ($currentSiteCount -gt 0 )
        {
            $SiteCollections = @() # array of objects
            foreach ($siteCollection in $db.Sites)
            {
                $url = $siteCollection.Url
                $SiteCollections += $url
                $siteCollection.Dispose()
            }
            
            $contentDatabaseProperties.SiteCollections = $SiteCollections
        }
        
        $ClonedContentDatabase = New-Object PSObject -Property $contentDatabaseProperties
        $ContentDatabasesProperty +=  $ClonedContentDatabase 
    }
    return $ContentDatabasesProperty
}

function GetWebApplicationFeatureDefinitionsClonedInstances([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    if ( $webApplication -eq $null) {return $null}
    if ( $webApplication.Id -eq $null) {return $null}
    
    # Check that the ID property gives back a valid GUID
    $guid = $webApplication.Id -as [System.GUID]
    if ($guid -eq $null) {return $null}
    
    
    $WebApplicationFeaturesProperty = @() # array of objects
    $activatedFeatures = Get-SPFeature -Limit ALL -WebApplication (Get-SPWebApplication -Identity $webApplication.Id.ToString())
    $webApplicationFeatures = Get-SPFeature -Limit ALL | where { $_.Scope -eq "WebApplication"}
    foreach ($f in $webApplicationFeatures)
    {
        $featureProperties = @{} #Hash table
        $internalName = $f.DisplayName
        $featureProperties.Name = $f.DisplayName
        $featureProperties.Title = $f.GetTitle([System.Globalization.CultureInfo]::CurrentUICulture)
        $featureProperties.Description = $f.GetDescription([System.Globalization.CultureInfo]::CurrentUICulture)
        $featureProperties.GUID = $f.Id
        $featureProperties.Hidden = $f.Hidden
        $featureProperties.Version = "Not Set"
        if ( $featureProperties.Version -ne $null)
        {
            $featureProperties.Version = $f.Version.ToString()
        }
        $featureProperties.RootDirectory = $f.RootDirectory.Replace("$sharepointRoot" , "{SharePointRoot}\")
        $featureProperties.RequireResources = $f.RequireResources
        $featureProperties.DefaultResourceFile = $f.DefaultResourceFile
        $featureProperties.IsActivated = $false
        # Check if the feature is activated
        foreach ($activatedFeature in $activatedFeatures)
        {
            if ($activatedFeature.Id -eq $f.Id)
            {
                $featureProperties.IsActivated = $true
                break;
            }
        }
        
        $ClonedFeatureSettings = New-Object PSObject -Property $featureProperties
        $WebApplicationFeaturesProperty +=  $ClonedFeatureSettings 
    }
    return $WebApplicationFeaturesProperty
}


function TryGetSPSite([System.Guid]$guid)
{
    try {
        if ($guid -eq $null) {return $null}
        $spSite = Get-SPSite -Identity $guid.ToString() -ErrorAction SilentlyContinue
        return $spSite
    }
    catch [Exception] {
         #$msg = $_.Exception.Message
         return $null
    }
}


function GetSiteCollectionFeatureDefinitionsClonedInstances([Microsoft.SharePoint.SPSite]$siteCollection)
{

    if ( $siteCollection -eq $null) {return $null}
    if ( $siteCollection.ID -eq $null) {return $null}
    $guid = $siteCollection.ID -as [System.GUID]
    
    # try get the SPSite Object from its GUID
    $spSite = TryGetSPSite($guid)
    if ($spSite -eq $null) {return $null}
    
    $SiteCollectionFeaturesProperty = @() # array of objects
    $activatedFeatures = Get-SPFeature -Limit ALL -Site $spSite
    $siteCollectionFeatures = Get-SPFeature -Limit ALL | where { $_.Scope -eq "Site"}
    foreach ($f in $siteCollectionFeatures)
    {
        $featureProperties = @{} #Hash table
        $internalName = $f.DisplayName
        $featureProperties.Name = $f.DisplayName
        $featureProperties.Title = $f.GetTitle([System.Globalization.CultureInfo]::CurrentUICulture)
        $featureProperties.Description = $f.GetDescription([System.Globalization.CultureInfo]::CurrentUICulture)
        $featureProperties.GUID = $f.Id
        $featureProperties.Hidden = $f.Hidden
        $featureProperties.Version = "Not Set"
        if ( $featureProperties.Version -ne $null)
        {
            $featureProperties.Version = $f.Version.ToString()
        }
        $featureProperties.RootDirectory = $f.RootDirectory.Replace("$sharepointRoot" , "{SharePointRoot}\")
        $featureProperties.RequireResources = $f.RequireResources
        $featureProperties.DefaultResourceFile = $f.DefaultResourceFile
        $featureProperties.IsActivated = $false
        # Check if the feature is activated
        foreach ($activatedFeature in $activatedFeatures)
        {
            if ($activatedFeature.Id -eq $f.Id)
            {
                $featureProperties.IsActivated = $true
                break;
            }
        }
        
        $ClonedFeatureSettings = New-Object PSObject -Property $featureProperties
        $SiteCollectionFeaturesProperty +=  $ClonedFeatureSettings 
    }
    return $SiteCollectionFeaturesProperty
}


function GetQuotaTemplateName($quotaID)
{
    $quotaTemplateName = "No Quota Template has been applied to this Site Collection"
    foreach ($quotaTemplate in $quotaTemplates)
    {
        if ($quotaTemplate.QuotaID -eq $quotaID)
        {
            $quotaTemplateName = $quotaTemplate.Name
            break;
        }
    }
    return $quotaTemplateName
}


function GetWebTemplateTitleForRootSite([Microsoft.SharePoint.SPSite]$siteCollection) 
{
    $templateName = "Not Set"
    if ($siteCollection.RootWeb -eq $null )
    {
        return $templateName
    }
    
    if ($siteCollection.RootWeb.Language -eq $null )
    {
        return "Sorry, cannot access Web templates"
    }
    
    $lcid = $siteCollection.RootWeb.Language
    $webTemplateIdForRootSite = $siteCollection.RootWeb.Configuration
    $webTemplateForRootSite = $siteCollection.RootWeb.WebTemplate
    $webTemplateNameForRootSite = [string]::Format("{0}#{1}",$webTemplateForRootSite,$webTemplateIdForRootSite)
    
    $webTemplates = $siteCollection.GetWebTemplates($lcid)
    foreach ($webTemplate in $webTemplates)
    {
        if ( $webTemplate.Name.Equals($webTemplateNameForRootSite))
        {
            #$templateName = [string]::Format("{0}.{1}",$webTemplate.Title,$webTemplate.Description)
            $templateName = [string]::Format("{0}",$webTemplate.Title)
            break;
        }
    }
    
    return $templateName
}

function GetRawUserSolutions([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access Solution Gallery"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    try {
        $url = $siteCollection.Url
        $userSolutions = Get-SPUserSolution -Site "$url" -ErrorAction SilentlyContinue
        if ( $userSolutions -eq $null )
        {
            return $errMsg
        }
        
        $count = 0
        $result = [String]::Empty
        foreach ($userSolution in $userSolutions)
        {
            $name = $userSolution.Name
            $status = SafeToString($userSolution.Status)
            
            if ($count -gt 0)
            {
                $result += ", "
            }
            $result += [string]::Format("{0} ({1})",$name,$status)
            
            $count += 1
        }
        return $result
    }
    catch [Exception] {
         #$msg = $_.Exception.Message
         return $errMsg
    }
    
    return $null
}

function GetUserActivityOnRootWebForTheCurrentMonth
{
    param([Microsoft.SharePoint.SPSite]$siteCollection, [String]$userName)
    
    if ($siteCollection -eq $null)
    {
        return $null
    }
    
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $null
        }
        
        $usageReportType = [Microsoft.SharePoint.Administration.SPUsageReportType]::user
        $usagePeriodType = [Microsoft.SharePoint.Administration.SPUsagePeriodType]::lastMonth
        $now = [System.DateTime]::Now
        [System.Data.DataTable]$dt = $rootweb.GetUsageData($usageReportType,$usagePeriodType , 1,$now)
        
        $reportedUserName = [String]::Empty
        $reportedNumberOfAccess = 0
        foreach ($dataRow in $dt.Rows)
        {
            $value = $dataRow[0]
            if ( $value.Contains($userName))
            {
                $reportedUserName = $userName
                $reportedNumberOfAccess = $dataRow[2]
                break;
            }
        }
        
        if ( $reportedUserName.Equals([String]::Empty))
        {
            return $null
        }
        
        return $reportedNumberOfAccess
        
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         return $null
    }
    
    return $null
}

function GetRawUsers([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access Users Info"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    try {
        $url = $siteCollection.Url
        $users = Get-SPUser -Web "$url" -ErrorAction SilentlyContinue
        if ( $users -eq $null )
        {
            return $errMsg
        }
        
        $count = 0
        $result = [String]::Empty
        foreach ($user in $users)
        {
            $groups = $user.Groups
            if ($groups -eq $null )
            {
                continue;
            }
            
            if ($groups.Count -eq 0 )
            {
                continue;
            }
            
            $name = $user.DisplayName
            $isSiteAdmin = SafeToString($user.IsSiteAdmin)
            $isSiteAuditor = SafeToString($user.IsSiteAuditor)
            
            if ($count -gt 0)
            {
                $result += ", "
            }
            
            $userInfo = [string]::Format("{0}",$name)
            
            
            if ( $isSiteAdmin -eq $true)
            {
                $userInfo = [string]::Format("{0} (SiteAdmin)",$name)
            }
            
            if ( $isSiteAuditor -eq $true)
            {
                if ( $isSiteAdmin -eq $false)
                {
                    $userInfo = [string]::Format("{0} (SiteAuditor)",$name)
                }
            }
            
            $result += $userInfo
            
            $count += 1
        }
        
        if ($result.Equals([String]::Empty))
        {
            $result = "No registered users"
        }
        
        return $result
    }
    catch [Exception] {
         #$msg = $_.Exception.Message
         return $errMsg
    }
    
    return $null
}


function GetUsersActivityForTheCurrentMonth([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access Users Info"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    try {
        $url = $siteCollection.Url
        $users = Get-SPUser -Web "$url" -ErrorAction SilentlyContinue
        if ( $users -eq $null )
        {
            return $errMsg
        }
        
        $count = 0
        $result = [String]::Empty
        foreach ($user in $users)
        {
            $groups = $user.Groups
            if ($groups -eq $null )
            {
                continue;
            }
            
            if ($groups.Count -eq 0 )
            {
                continue;
            }
            
            $name = $user.DisplayName
            $activity = GetUserActivityOnRootWebForTheCurrentMonth $siteCollection $name
            
            if ($activity -eq $null) { continue}
            
            if ($count -gt 0)
            {
                $result += ", "
            }
            
            $userInfo = [string]::Format("{0} (#access={1})",$name, $activity)
            $result += $userInfo
            
            $count += 1
        }
        
        if ($result.Equals([String]::Empty))
        {
            $result = "No available usage data"
        }
        
        return $result
    }
    catch [Exception] {
         #$msg = $_.Exception.Message
         return $errMsg
    }
    
    return $null
}

function GetRootWebInfos([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $webTemplate = $rootWeb.WebTemplate
        $webTemplateId = $rootWeb.Configuration
        $title = $rootWeb.Title 
        $templateTitle = GetWebTemplateTitleForRootSite($siteCollection)
        $rootWebInfo = [string]::Format("Title= '{0}' (Template = {1}#{2} => {3})",$title,$webTemplate,$webTemplateId, $templateTitle)
        
        return $rootWebInfo
        
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         return $errMsg
    }
    
    return $null
}


function GetRootWebLanguage([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $lcid = $rootWeb.Language -as [System.Int32]
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo($lcid)
        $languageName = $culture.EnglishName
        $isMultilingual = $rootWeb.IsMultilingual
        $result = [string]::Format("'{0}' (Multilingual is enabled = {1})",$languageName,$isMultilingual)
        
        return $result
        
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         return $errMsg
    }
    
    return $null
    
}


function GetWebsAndListsWithUniquePermission([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $infos = $rootWeb.GetWebsAndListsWithUniquePermissions()
        $noUniquePermissionMsg = "Perm. inheritance has not been broken on this Site Collection"
        if ($infos -eq  $null)
        {
            return $noUniquePermissionMsg
        }
        
        if ($infos.Count -eq  0)
        {
            return $noUniquePermissionMsg
        }
        
        $count = 0
        $result = [String]::Empty
        foreach ($info in $infos)
        {
            $url = SafeToString($info.Url)
            
            if ($url -eq $null) { continue }
            if ($url.Equals([String]::Empty)) { continue }
            
            # Check if current $info represents the root web
            if ($rootWeb.ServerRelativeUrl.Contains($url)) { continue }
            
            # Hide IWConvertedForms
            if ($url.Contains("IWConvertedForms")) { continue }
            
            # Hide ContentTypeSyncLog
            if ($url.Contains("ContentTypeSyncLog")) { continue }
            
            # Hide _catalogs/users
            if ($url.Contains("_catalogs/users")) { continue }
            
            # Hide TaxonomyHiddenList
            if ($url.Contains("TaxonomyHiddenList")) { continue }
            
            # Hide ContentTypeAppLog
            if ($url.Contains("ContentTypeAppLog")) { continue }
            
            # Hide PackageList 
            if ($url.Contains("PackageList")) { continue }
            
            $url = "/" + $url
            $rootWebRelativeUrl = $rootWeb.ServerRelativeUrl
            
            # Simplify url (if needed)
            if ($url.Contains($rootWebRelativeUrl)) {
                $url = $url.Replace($rootWebRelativeUrl,[String]::Empty)
            }
            
            if ($count -gt 0)
            {
                $result += ", "
            }
            
            $objectType = SafeToString($info.Type)

            $webListInfo = [string]::Format("{0} ({1})",$url,$objectType)
            
            $result += $webListInfo
            
            $count += 1
        } # end foreach ($info in $infos)

        if ( $result.Equals([String]::Empty))
        {
            return $noUniquePermissionMsg
        }
        
        return $result
        
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         return $errMsg
    }
    
    return $null
    
}


function TryGetSiteCollectionsFromContentDatabase([Microsoft.SharePoint.Administration.SPContentDatabase]$contentDb)
{
    try {  
        $sites = $contentDb.Sites
        if ( $sites -eq $null ) { return $null}
        if ( $sites.Count -eq 0 ) { return $null}    
        return $sites
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         return $null
    }
    
    return $null
}

function TryFindContentDatabase([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot find Content Database"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    $guid = $siteCollection.ID
    
    if ($guid -eq $null)
    {
        return $errMsg
    }
    $failedContentDbs =  @() # array of objects   
    try {
        $dbs = Get-SPDatabase
        foreach ($db in $dbs)
        {
            $contentDb = $db -as [Microsoft.SharePoint.Administration.SPContentDatabase]
            if ( $contentDb -eq $null ) { continue}
            
            if ($contentDb.Name -eq $null) {continue}
            
            $sites = TryGetSiteCollectionsFromContentDatabase($contentDb)
            if ($sites -eq $null)
            {
                $failedContentDbs += $contentDb.Name
                continue;
            }
            
            foreach ($site in $sites)
            {
                if ($site.ID -eq $null) { continue}
                if ($site.ID.Equals($guid))
                {
                    return $contentDb
                }
            } # end foreach ($site in $sites)
        } # end foreach ($db in $dbs)
        
        if ( $failedContentDbs -eq $null) {return $errMsg}
        if ( $failedContentDbs.Count -eq 0) {return $errMsg}
        
        $info = [string]::Empty
        $count = 0
        foreach ($db in $failedContentDbs)
        {
            if ($count -gt 0)
            {
                $info += ", "
            }
            $info += SafeToString($db)
            $count += 1
        }
        
        $infos = [string]::Format("Sorry, cannot find Content Database : maybe the following database(s) are not mounted : {0}",$info)
        
        return $infos
        
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         return $errMsg
    }
    
    return $null
}


function GetContentDatabaseInfos([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access Content Database"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    try {
        $db = $siteCollection.ContentDatabase
        if ($db -eq $null)
        {
            # try to find why this content database is missing
            $infos = TryFindContentDatabase($siteCollection)
            if ( $infos -eq $null ) {return $errMsg}
            $errMsg2 = $infos -as [String]
            if ( $errMsg2 -ne $null ) {return $errMsg2}
            return $errMsg
        }
        
        $dbName = SafeToString($db.Name)
        $dbServer = SafeToString($db.Server)
        
        $infos = [string]::Format("{0} (On server {1})",$dbName,$dbServer)
        
        return $infos
        
    }
    catch [Exception] {
         #$siteCollectionProperties.Owner = $_.Exception.Message
         $infos = TryFindContentDatabase($siteCollection)
         if ( $infos -eq $null ) {return $errMsg}
         $errMsg2 = $infos -as [String]
         if ( $errMsg2 -ne $null ) {return $errMsg2}
         return $errMsg
    }
    
    return $null
}


function GetSiteCollectionCreationDate([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $date = $rootWeb.Created -as [System.DateTime]
        if ( $date -eq $null ) 
        {
            return $null
        }
        
        #Convert the creation date (which is UTC ) to a local datetime
        $date = $date.ToLocalTime()
        $dateAsString = SafeToString($date)
        
        return $dateAsString
        
    }
    catch [Exception] {
         #Write-Host $_.Exception.Message
         return $errMsg
    }
    
    return $null
    
}

function GetPermissionLevels([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $permissionLevels = $rootWeb.RoleDefinitions
 
        if ( $permissionLevels -eq $null ) 
        {
            return $null
        }
        
        if ( $permissionLevels.Count -eq 0 ) 
        {
            return $null
        }
        
        $SiteCollectionRolesProperty = @() # array of objects
        
        foreach ($permissionLevel in $permissionLevels)
        {
            $permissionLevelProperties = @{} #Hash table
            $permissionLevelProperties.Name = $permissionLevel.Name
            $permissionLevelProperties.Description = $permissionLevel.Description
            $permissionLevelProperties.Type = SafeToString($permissionLevel.Type)
            
            $ClonedPermissionLevelSettings = New-Object PSObject -Property $permissionLevelProperties
            $SiteCollectionRolesProperty +=  $ClonedPermissionLevelSettings 
        }
        return $SiteCollectionRolesProperty
        
    }
    catch [Exception] {
         #Write-Host $_.Exception.Message
         return $errMsg
    }
    
    return $null
}

function GetRoleDefinitions([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $permissionLevels = $rootWeb.RoleDefinitions
 
        if ( $permissionLevels -eq $null ) 
        {
            return $null
        }
        
        if ( $permissionLevels.Count -eq 0 ) 
        {
            return $null
        }
        
        $SiteCollectionRolesProperty = @() # array of objects
        
        foreach ($permissionLevel in $permissionLevels)
        {
            $permissionLevelProperties = @{} #Hash table
            $permissionLevelProperties.Name = $permissionLevel.Name
            $ClonedPermissionLevelSettings = New-Object PSObject -Property $permissionLevelProperties
            $SiteCollectionRolesProperty +=  $ClonedPermissionLevelSettings 
        }
        
        # add a special entry for the anonymous access rights
        $permissionLevelProperties = @{} #Hash table
        $permissionLevelProperties.Name = "Anonymous Users"
        $ClonedPermissionLevelSettings = New-Object PSObject -Property $permissionLevelProperties
        $SiteCollectionRolesProperty +=  $ClonedPermissionLevelSettings 
    
        return $SiteCollectionRolesProperty
        
    }
    catch [Exception] {
         #Write-Host $_.Exception.Message
         return $errMsg
    }
    
    return $null
}


function GetFullBasePermissionsForAllRoles([Microsoft.SharePoint.SPSite]$siteCollection)
{
    $errMsg = "Sorry, cannot access RootWeb"
    
    if ($siteCollection -eq $null)
    {
        return $errMsg
    }
    
    try {
        $rootWeb = $siteCollection.RootWeb
        if ($rootWeb -eq $null)
        {
            return $errMsg
        }
        
        $permissionLevels = $rootWeb.RoleDefinitions
 
        if ( $permissionLevels -eq $null ) 
        {
            return $null
        }
        
        if ( $permissionLevels.Count -eq 0 ) 
        {
            return $null
        }
        
        #
        $BasePermissionsCollectionProperty = @() # array of objects
        
        $enumValues = [System.Enum]::GetValues([Microsoft.SharePoint.SPBasePermissions])
        foreach ($permissionValue in $enumValues)
        {
            $basePermissionProperties = @{} #Hash table
            $basePermissionProperties.Name = $permissionValue.ToString()
            $basePermissionProperties.Description = GetBasePermissionDescription($permissionValue)
            $emptymask = [Microsoft.SharePoint.SPBasePermissions]::EmptyMask
            $fullMask = [Microsoft.SharePoint.SPBasePermissions]::FullMask
            
            #do not record special value like EmptyMask
            if ($permissionValue -eq $emptymask)
            {
                continue
            }
            
            #do not record special value like FullMask
            if ($permissionValue -eq $fullMask)
            {
                continue
            }
            
            $basePermissionProperties.RoleDefinitions = @() # array of objects
            # Check if $permissionValue is granted within the available Permission Levels
            foreach ($roleDefinition in $permissionLevels)
            {
                $roleDefinitionProperties = @{} #Hash table
                $roleDefinitionProperties.Name = $roleDefinition.Name
                $basePermission = $roleDefinition.BasePermissions
                
                if ($basePermission -eq $emptymask)
                {
                    $roleDefinitionProperties.Enabled = $false
                    $ClonedRoleDefinitionSettings = New-Object PSObject -Property $roleDefinitionProperties
                    $basePermissionProperties.RoleDefinitions +=  $ClonedRoleDefinitionSettings
                    continue
                }
                if ($basePermission -eq $fullMask)
                {
                    $roleDefinitionProperties.Enabled = $true
                    $ClonedRoleDefinitionSettings = New-Object PSObject -Property $roleDefinitionProperties
                    $basePermissionProperties.RoleDefinitions +=  $ClonedRoleDefinitionSettings
                    continue
                }
                
                $roleDefinitionProperties.Enabled = $false
                if ( ($permissionValue -band $basePermission) -eq $permissionValue)
                {
                    $roleDefinitionProperties.Enabled = $true
                }
                
                $ClonedRoleDefinitionSettings = New-Object PSObject -Property $roleDefinitionProperties
                $basePermissionProperties.RoleDefinitions +=  $ClonedRoleDefinitionSettings 
            } # end foreach ($roleDefinition in $permissionLevels)
            
            # add a special entry for anonymous access rights
                $roleDefinitionProperties = @{} #Hash table
                $roleDefinitionProperties.Name = "Anonymous Users"
                $basePermission = $rootWeb.AnonymousPermMask64
                
                $roleDefinitionProperties.Enabled = $false
                if ( ($permissionValue -band $basePermission) -eq $permissionValue)
                {
                    $roleDefinitionProperties.Enabled = $true
                }
                
                if ($basePermission -eq $emptymask)
                {
                    $roleDefinitionProperties.Enabled = $false
                }
                
                if ($basePermission -eq $fullMask)
                {
                    $roleDefinitionProperties.Enabled = $true
                }
                
                $ClonedRoleDefinitionSettings = New-Object PSObject -Property $roleDefinitionProperties
                $basePermissionProperties.RoleDefinitions +=  $ClonedRoleDefinitionSettings 
            # end special entry for anonymous access rights
            
            
            $ClonedBasePermissionSettings = New-Object PSObject -Property $basePermissionProperties
            $BasePermissionsCollectionProperty +=  $ClonedBasePermissionSettings 
            
        } #foreach ($permissionValue in $enumValues)

        return $BasePermissionsCollectionProperty
        
    } # end try
    catch [Exception] {
         #Write-Host $_.Exception.Message
         return $errMsg
    }
    
    return $null
    
}


function GetWebApplicationSiteCollections($webApplication)
{
    $GC = Start-SPAssignment
    $siteCollections = $GC | Get-SPSite -Limit ALL -WebApplication (Get-SPWebApplication -Identity $webApplication.Id.ToString())
    
    $SiteCollectionsProperty = @() # array of objects
    foreach ($siteCollection in $siteCollections)
    {
        $siteCollectionProperties = @{} #Hash table
        $siteCollectionProperties.ID = $siteCollection.ID
        $siteCollectionProperties.Url = $siteCollection.Url
        #$siteCollectionProperties.AllowUnsafeUpdates = $siteCollection.AllowUnsafeUpdates
        #$siteCollectionProperties.LetSharePointTrapAndHandleAccessDeniedExceptions = $siteCollection.CatchAccessDeniedException
        $siteCollectionProperties.UnavailableForReadAccess = $siteCollection.ReadLocked
        $siteCollectionProperties.ReadOnly = $siteCollection.ReadOnly
        $siteCollectionProperties.UnavailableForWriteAccess = $siteCollection.WriteLocked
        
        # Get Site Collection Creation Date 
        $siteCollectionProperties.CreationDate = GetSiteCollectionCreationDate($siteCollection)
        
        # Get Site Collection Owner
        try {
            $siteCollectionProperties.Owner = "Not Set"
            if ($siteCollection.Owner -ne $null)
            {
                $siteCollectionProperties.Owner = $siteCollection.Owner.ToString()
            }
        }
        catch [Exception] {
             $siteCollectionProperties.Owner = $_.Exception.Message
        }

        # try get Content Database
        $siteCollectionProperties.ContentDatabase = GetContentDatabaseInfos($siteCollection)
      
        # try to get the number of Web Sites inside the Site Collection
        try {
            $siteCollectionProperties.NumberOfWebSites = 0
            if ($siteCollection.AllWebs -ne $null )
            {
                $siteCollectionProperties.NumberOfWebSites = $siteCollection.AllWebs.Count
            }
        }
        catch [Exception] {
             #$siteCollectionProperties.Owner = $_.Exception.Message
             $siteCollectionProperties.NumberOfWebSites = $null
        }
        
        # try get Quota Template associated to this site collection
        $siteCollectionProperties.Quota = $null
        if ($siteCollection.Quota -ne $null )
        {
            $siteCollectionProperties.Quota = GetQuotaTemplateName($siteCollection.Quota.QuotaID)
        }
        
        # try to get the disk space usage 
        try {
            
            if ($siteCollection.Usage -ne $null )
            {
                $storage = $siteCollection.Usage.Storage
                $storageInGB = ([Double]$storage)/[Double](1024*1024*1024)
                $storageMaxInGB = "Unknown max storage value"
                $storageMax = -1
                # try get Quota for disk usage
                if ($siteCollection.Quota -ne $null )
                {
                    $storageMax = $siteCollection.Quota.StorageMaximumLevel
                    $storageMaxInGB = ([Double]$storageMax)/[Double](1024*1024*1024)
                }
                
                $siteCollectionProperties.DiskStorageUsed = [string]::Format("{0:F2} Gb (max = {1:F2} Gb)",$storageInGB,$storageMaxInGB) 
                
                if ($storageMax -eq  0 )
                {
                    $siteCollectionProperties.DiskStorageUsed = [string]::Format("{0:F2} Gb (max = unlimited)",$storageInGB) 
                } 
            }
        }
        catch [Exception] {
             #$siteCollectionProperties.Owner = $_.Exception.Message
             $siteCollectionProperties.DiskStorageUsed = $null
        }
        
        
        # try to get resource usage for User Solutions
        try {
            $currentResourceUsage = "Unknown"
            $averageResourceUsage = "Unknown"
            $maxResourceUsage = "Unknown"
            $siteCollectionProperties.ResourceQuotaExceeded = $false
            
            $siteCollectionProperties.ComputerResourceUsed = [string]::Format("{0} points (max = {1} points / Average = {2} points)",$currentResourceUsage,$maxResourceUsage,$averageResourceUsage) 
                
            if ($siteCollection.CurrentResourceUsage -ne $null )
            {
                $currentResourceUsage = $siteCollection.CurrentResourceUsage
                $averageResourceUsage = $siteCollection.AverageResourceUsage
                $siteCollectionProperties.ResourceQuotaExceeded = $siteCollection.ResourceQuotaExceeded
             
                # try get Quota for Resource
                if ($siteCollection.Quota -ne $null )
                {
                    $maxResourceUsage = $siteCollection.Quota.UserCodeMaximumLevel
                }
                
                $siteCollectionProperties.ComputerResourceUsed = [string]::Format("{0:F2} points (max = {1:F2} points / Average = {2:F2} points)",$currentResourceUsage,$maxResourceUsage,$averageResourceUsage) 
                
                if ($maxResourceUsage -eq  0 )
                {
                    $siteCollectionProperties.ComputerResourceUsed = [string]::Format("No Computer Resource used because UserSolutions are disabled") 
                } 
            }
        }
        catch [Exception] {
             #$siteCollectionProperties.Owner = $_.Exception.Message
             $siteCollectionProperties.ComputerResourceUsed = $null
        }
        
        # try to get root web infos 
        $siteCollectionProperties.RootWeb = GetRootWebInfos($siteCollection)
        
        # try to get root web default language 
        $siteCollectionProperties.Language = GetRootWebLanguage($siteCollection)

        #try get User Solutions
        $siteCollectionProperties.UserSolutions = GetRawUserSolutions($siteCollection)
        
        #try get Users
        $siteCollectionProperties.Users = GetRawUsers($siteCollection)
        
        #try get Users Activity for the current Month
        $siteCollectionProperties.UsersActivityForTheCurrentMonth = GetUsersActivityForTheCurrentMonth($siteCollection)
        
        #try get Webs and Lists with unique Permissions
        $siteCollectionProperties.WebsAndListsWithUniquePermission =GetWebsAndListsWithUniquePermission($siteCollection)
          
        #
        $SiteCollectionSettings = New-Object PSObject -Property $siteCollectionProperties
        $SiteCollectionsProperty +=  $SiteCollectionSettings 
    }
    
    Stop-SPAssignment $GC
    return $SiteCollectionsProperty
}


function GetSecuritySettingsForAnonymousUsers([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $anonymousPolicy = $webApplication.Policies.AnonymousPolicy
    #
    $UserPolicyProperties = @{} #Hash table
    $UserPolicyProperties.DisplayName="Anonymous Users"
    $UserPolicyProperties.Authorizations = SafeToString($anonymousPolicy)
    #
    $UserPolicySettings = New-Object PSObject -Property $UserPolicyProperties
    
    return $UserPolicySettings
}

function GetSecuritySettingsForSearchAndOtherUsers([Microsoft.SharePoint.Administration.SPWebApplication]$webApplication)
{
    $UserPolicyCollectionProperty = @() # array of objects
    $userPolicies = $webApplication.Policies
    foreach ($userPolicy in $userPolicies)
    {
        $UserPolicyProperties = @{} #Hash table
        $UserPolicyProperties.Name = $userPolicy.UserName
        $UserPolicyProperties.DisplayName = $userPolicy.DisplayName
        $UserPolicyProperties.IsSiteCollectionAdministrator = $false
        $UserPolicyProperties.IsSiteCollectionAuditor = $false
        
        # Get every authorizations for the current user  
        $UserPolicyProperties.Authorizations = [String]::Empty
        $count = 0
        foreach ($policyRole in $userPolicy.PolicyRoleBindings)
        {
            if ($policyRole.IsSiteAdmin -eq $true)
            {
                $UserPolicyProperties.IsSiteCollectionAdministrator = $true
            }
            if ($policyRole.IsSiteAuditor -eq $true)
            {
                $UserPolicyProperties.IsSiteCollectionAuditor = $true
            }
            if ($count -gt 0)
            {
                $UserPolicyProperties.Authorizations += ", "
            }
            $UserPolicyProperties.Authorizations += $policyRole.Name
            $count += 1
        } # end foreach ($policyRole in $userPolicy.PolicyRoleBindings)
        
        $UserPolicySettings = New-Object PSObject -Property $UserPolicyProperties
        $UserPolicyCollectionProperty +=  $UserPolicySettings 
    } # end foreach ($userPolicy in $userPolicies)
    
    # Get Anonymous Policy
    $UserPolicyCollectionProperty += GetSecuritySettingsForAnonymousUsers($webApplication)
    
    return $UserPolicyCollectionProperty
}


function GetAllSiteCollections()
{
    $GC = Start-SPAssignment
    $siteCollections = $GC | Get-SPSite -Limit ALL
    
    $SiteCollectionsProperty = @() # array of objects
    foreach ($siteCollection in $siteCollections)
    {
        $siteCollectionProperties = @{} #Hash table
        $siteCollectionProperties.ID = $siteCollection.ID
        $siteCollectionProperties.Url = $siteCollection.Url
        $siteCollectionProperties.SiteCollectionFeatures = GetSiteCollectionFeatureDefinitionsClonedInstances($siteCollection)
        $siteCollectionProperties.SiteCollectionFeaturesCount = GetArrayItemCount($siteCollectionProperties.SiteCollectionFeatures)
        $siteCollectionProperties.PermissionLevels = GetPermissionLevels($siteCollection)
        $siteCollectionProperties.PermissionLevelsCount = GetArrayItemCount($siteCollectionProperties.PermissionLevels)
        $siteCollectionProperties.RoleDefinitions = GetRoleDefinitions($siteCollection)
        $siteCollectionProperties.PermissionLevelsMatrix = GetFullBasePermissionsForAllRoles($siteCollection)
        #
        $SiteCollectionSettings = New-Object PSObject -Property $siteCollectionProperties
        $SiteCollectionsProperty +=  $SiteCollectionSettings 
    }
    
    Stop-SPAssignment $GC
    return $SiteCollectionsProperty
}


WaitFor2Seconds
# End activity

###########################################################
# Begin activity
WriteProgress "Reading Farm properties ..." 10
$SPFarm = GetLocalFarm

$properties = @{}
$properties.BuildVersion = $SPFarm.BuildVersion.ToString()
$properties.Version = $SPFarm.Version.ToString()
$properties.CEIPEnabled = $SPFarm.CEIPEnabled
$properties.DaysBeforePasswordExpirationToSendEmail = $SPFarm.DaysBeforePasswordExpirationToSendEmail
$properties.CEIPEnabled = $SPFarm.CEIPEnabled
$properties.DefaultServiceAccount = $SPFarm.DefaultServiceAccount.LookupName()
$properties.DisplayName = $SPFarm.DisplayName
$properties.Status = $SPFarm.Status
$properties.Properties = New-Object PSObject -Property $SPFarm.Properties
$properties.ConfigurationDatabaseConnectionString = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\Secure\ConfigDB').dsn
$properties.ConfigurationDatabaseID = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\Secure\ConfigDB').id
$properties.ServerLanguage = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\ServerLanguage')
$properties.CentralAdminUrl = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\WSS').CentralAdministrationURL
$centralAdminUrl = $properties.CentralAdminUrl
$properties.SharePointRoot = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0').Location
$sharepointRoot = $properties.SharePointRoot
$properties.InstalledLanguages = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\InstalledLanguages')

$installedProducts = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\14.0\WSS\InstalledProducts')
foreach ($p in $installedProducts)
{
    if ( $p -match ".*BEED1F75.*") {$properties.LicenceType = "SharePoint Foundation"}
    if ( $p -match ".*1328E89E.*") {$properties.LicenceType += "; Search Server Express 2010"}
    if ( $p -match ".*B2C0B444.*") {$properties.LicenceType += "; SharePointServer2010StandardTrial"}
    if ( $p -match ".*3FDFBCC8.*") {$properties.LicenceType += "; SharePointServer2010Standard"}
    if ( $p -match ".*88BED06D.*") {$properties.LicenceType += "; SharePoint Server 2010 Enterprise Trial"}
    if ( $p -match ".*D5595F62.*") {$properties.LicenceType += "; SharePoint Server 2010 Enterprise"}
    if ( $p -match ".*BC4C1C97.*") {$properties.LicenceType += "; Search Server 2010 Trial"}
    if ( $p -match ".*08460AA2.*") {$properties.LicenceType += "; Search Server 2010"}
    if ( $p -match ".*84902853.*") {$properties.LicenceType += "; Project Server 2010 Trial"}
    if ( $p -match ".*ED21638F.*") {$properties.LicenceType += "; Project Server 2010"}
    if ( $p -match ".*926E4E17.*") {$properties.LicenceType += "; Office Web Applications 2010"}
}

$properties.WebApplicationsList = GetAllWebApplications

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Getting SPWebService layer ..." 20
$spWebService = GetSPWebServiceInstance
#
$properties.ActiveDirectoryDomain = $spWebService.ActiveDirectoryDomain
$properties.BrowserCEIPEnabled = $spWebService.BrowserCEIPEnabled
#
WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Reading Service Application Proxy Groups ..." 25
$proxyGroups = Get-SPServiceApplicationProxyGroup
$properties.ServiceApplicationProxyGroups = @() # array of objects
$properties.ServiceApplicationProxyGroupsCount = 0

foreach ($proxyGroup in $proxyGroups)
{
    $proxyGroupProperties = @{} #Hash table
    $proxyGroupProperties.Name = $proxyGroup.FriendlyName
    $proxyGroupProperties.ServiceApplicationProxies = @() # array of objects
    # get all proxies in this group
    $proxyGroupProperties.ServiceApplicationProxiesCount = 0
    foreach ($proxy in $proxyGroup.Proxies)
    {
        $ClonedProxy = GetProxyClonedInstance($proxy)
        $proxyGroupProperties.ServiceApplicationProxies +=  $ClonedProxy
        $proxyGroupProperties.ServiceApplicationProxiesCount += 1
    }
    
    $ClonedProxyGroup = New-Object PSObject -Property $proxyGroupProperties
    $properties.ServiceApplicationProxyGroups +=  $ClonedProxyGroup
    $properties.ServiceApplicationProxyGroupsCount += 1 
}

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Reading Quota Templates ..."  30
$spWebService = GetSPWebServiceInstance
$quotaTemplates = $spWebService.QuotaTemplates
$properties.QuotaTemplates = @() # array of objects
$properties.QuotaTemplatesCount = 0
foreach ($quotaTemplate in $quotaTemplates)
{
    $ClonedQuotaTemplate = GetQuotaTemplateClonedInstance($quotaTemplate)
    $properties.QuotaTemplates +=  $ClonedQuotaTemplate
    $properties.QuotaTemplatesCount += 1
}
#
WaitFor2Seconds
# End activity

###########################################################
# Begin activity
WriteProgress "Reading Farm Features ..." 40
$farmFeatures = Get-SPFeature -Limit ALL | where { $_.Scope -eq "Farm"}
$properties.FarmFeatures = @() # array of objects
$properties.FarmFeaturesCount = 0
foreach ( $f in $farmFeatures)
{
    $ClonedFeatureDefinition = GetFeatureDefinitionClonedInstance($f)
    $properties.FarmFeatures +=  $ClonedFeatureDefinition
    $properties.FarmFeaturesCount += 1
}

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Reading Developer DashBoard Settings ..." 45

$ClonedDeveloperDashboardSettings = GetDeveloperDashboardSettingsClonedInstance
$properties.DeveloperDashBoardSettings = $ClonedDeveloperDashboardSettings

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Reading Log Settings ..." 50

$ClonedLogSettings = GetDiagnosticsServiceClonedInstance
$properties.LogSettings = $ClonedLogSettings

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Reading deployed Farm Solutions ..."  53
$farmSolutions = Get-SPSolution

$properties.FarmSolutions = @() # array of objects
$properties.FarmSolutionsCount = 0
foreach ($solution in  $farmSolutions)
{
    $ClonedSolution = GetSolutionClonedInstance($solution)
    $properties.FarmSolutions += $ClonedSolution
    $properties.FarmSolutionsCount += 1
}

WaitFor2Seconds
# End activity



###########################################################
# Begin activity
WriteProgress "Reading Managed Account Settings ..." 55
$managedAccounts = Get-SPManagedAccount
#$managedAccounts = New-Object Microsoft.SharePoint.Administration.SPFarmManagedAccountCollection($SPFarm)

$properties.ManagedAccounts = @() # array of objects
$properties.ManagedAccountsCount = 0

foreach ( $account in $managedAccounts)
{
    $ClonedManagedAccount = GetManagedAccountClonedInstance($account)
    $properties.ManagedAccounts +=  $ClonedManagedAccount 
    $properties.ManagedAccountsCount += 1 
} # end foreach ( $account in $managedAccounts)

# Add LocalSystem
$LocalSystemAccount = GetManagedAccountClonedInstance("LocalSystem")
$properties.ManagedAccounts +=  $LocalSystemAccount  
# Add LocalService
$LocalServiceAccount = GetManagedAccountClonedInstance("LocalService")
$properties.ManagedAccounts +=  $LocalServiceAccount  
# Add NetworkService
$NetworkServiceAccount = GetManagedAccountClonedInstance("NetworkService")
$properties.ManagedAccounts +=  $NetworkServiceAccount  

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress  "Reading Web Applications (may take several minutes) ..." 60
$webApplications = Get-SPWebApplication -IncludeCentralAdministration
$properties.WebApplications = @() # array of objects
$properties.WebApplicationsCount = 0

foreach ( $webApplication in $webApplications)
{
    $webAppProperties = @{} #Hash table
    $webAppProperties.Name = $webApplication.DisplayName
    $webAppProperties.GlobalSettings = GetWebApplicationGlobalSettings($webApplication)
    $webAppProperties.InlineDownloadedMimeTypesSettings = GetInlineDownloadedMimeTypesSettings($webApplication)
    $webAppProperties.BlockedFileExtensionsSettings = GetBlockedFileExtensionsSettings($webApplication)
    $webAppProperties.ManagedPaths = GetManagedPathSettings($webApplication)
    $webAppProperties.LargeListThrottlingSettings = GetLargeListThrottlingSettings($webApplication)
    $webAppProperties.IISSettings = GetIISSettings($webApplication)
    $webAppProperties.SharePointDesignerSettings = GetSharePointDesignerSettings($webApplication)
    $webAppProperties.WebPartSettings = GetWebPartSettings($webApplication)
    $webAppProperties.PermissionsSettings = GetPermissionsSettings($webApplication)
    $webAppProperties.SecurityPolicies = GetSecuritySettingsForSearchAndOtherUsers($webApplication)
    $webAppProperties.SecurityPoliciesCount = GetArrayItemCount($webAppProperties.SecurityPolicies)
    $webAppProperties.ContentDatabasesCount = GetArrayItemCount($webApplication.ContentDatabases)
    if ( $webApplication.ContentDatabases.Count -eq 1) 
    {
        $webAppProperties.ContentDatabaseSettings = GetContentDatabasesSettings($webApplication)
    }
    if ( $webApplication.ContentDatabases.Count -gt 1) 
    {
        $webAppProperties.ContentDatabases = GetContentDatabasesSettings($webApplication)
    }
    $webAppProperties.WebApplicationFeatures = GetWebApplicationFeatureDefinitionsClonedInstances($webApplication)
    $webAppProperties.WebApplicationFeaturesCount = GetArrayItemCount($webAppProperties.WebApplicationFeatures)
    
    $webAppProperties.WebApplicationSiteCollections = GetWebApplicationSiteCollections($webApplication)
    $webAppProperties.WebApplicationSiteCollectionsCount = GetArrayItemCount($webAppProperties.WebApplicationSiteCollections)
    
    $ClonedWebApplication = New-Object PSObject -Property $webAppProperties
    $properties.WebApplications +=  $ClonedWebApplication
    $properties.WebApplicationsCount += 1
}

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress "Reading All Site Collections (may take several minutes)..." 70

$properties.AllSiteCollections = GetAllSiteCollections
$properties.AllSiteCollectionsCount = GetArrayItemCount($properties.AllSiteCollections)

WaitFor2Seconds
# End activity

###########################################################
# Begin activity
WriteProgress "Storing Farm Object as XML (may take several minutes)..." 75
$ClonedFarm = New-Object PSObject -Property $properties 

$xmlFilePath = "$Home\SPClonedFarm.xml"
$xmlFilePath2 = "$Home\SPClonedFarm2.xml"
$ClonedFarm | Export-Clixml -depth 2  -Path $xmlFilePath 

# Remove the xmlns declaration from the generated XML file
$xml = [System.IO.File]::ReadAllText($xmlFilePath)
$xml = $xml -replace "xmlns=""http://schemas.microsoft.com/powershell/2004/04""", [System.String]::Empty
# Wait for the file system to be stable
[System.Threading.Thread]::Sleep(3000)
# Write back the XML file
[System.IO.File]::WriteAllText($xmlFilePath2,$xml)

WaitFor2Seconds
# End activity

###########################################################
# Begin activity
WriteProgress  "Generating XSL..." 70

#Setup the Xsl 
$xsl=@'
<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
      <html>
        <head>
          <title>SPFarm Poster v{currentScriptVersion}</title>
          <style type="text/css">
            body
            {
            width:100%;
            font-size: 100%;
            font-family: Verdana, Tahoma, Arial, "Helvetica Neue" , Helvetica, Sans-Serif;
            margin: 0;
            padding: 0;
            color: #232323;
            }

            .autocenter
            {
            width: 90%;
            margin-left: auto;
            margin-right: auto;
            }

            .right
            {
            float: right;
            }

            .left
            {
            float: left;
            }

            h1
            {
            font-family: "Helvetica Neue" ,Helvetica,Arial,sans-serif;
            color: #333;
            font-size: 100%;
            }

            h2
            {
            margin-left: 20px;
            margin-bottom: 5px;
            font-size: 90%;
            }

            .round
            {
            -moz-border-radius: 5px;
            -webkit-border-radius: 5px;
            }

            .wide
            {
            /*width: 100%;*/
            width: 15000px;
            }

            .clear
            {
            clear: both;
            }

            .featureHeight
            {
            height: 410px;
            }

            .contentDBHeight
            {
            height: 220px;
            }

            .siteCollectionHeight
            {
            min-height: 500px;
            }

            .quotaTemplateHeight
            {
            height: 120px;
            }

            .farmSolutionHeight
            {
            height: 180px;
            }

            .proxyHeight
            {
            height: 190px;
            }

            .permissionLevelHeight
            {
            height: 150px;
            }

            .commonWidth
            {
            width: 650 px;
            }

            .common2Width
            {
            width: 1000 px;
            }



            .smallWidth
            {
            width: 10%;
            }

            .verysmallWidth
            {
            width: 5%;
            }

            .mediumWidth
            {
            width: 20%;
            }

            .spFarm
            {
            background-color: #D1FFA4;
            border: thin solid #006600;
            }

            .bgColorLightGreen
            {
            background-color: #EBFFD7;
            }

            .spFarmFeatures
            {
            /*background-color: #A5FEE3;*/
            border: thin solid #02C489;
            }

            .property
            {
            font-weight: bold;
            font-style: oblique;
            font-variant: normal;
            color: #3366FF;
            }

            .indent
            {
            margin: 15px;
            padding: 0;
            }

            .smallerFont
            {
            font-size : 0.8em;
            }

            .backgroundRed
            {
            background-color: #FFAEAE;
            }

            .backgroundGreen
            {
            background-color: #D1FFA4;
            }

            .backgroundOrange
            {
            background-color: #FFE066;
            }

            .backgroundYellow
            {
            background-color: #FFD52D;
            }

            .backgroundGray
            {
            background-color: #C0C0C0;
            color: #FFFFFF;
            }

            .backgroundGrayRed
            {
            background-color: #DEB4B4;
            color: #FFFFFF;
            }

            .backgroundGrayGreen
            {
            background-color: #B9E6AE;
            color: #FFFFFF;
            }

            .greenFont
            {
            color: #00B700;
            }

            #footer
            {
            text-align: center;
            padding: 8px 0;
            margin-top: .7em;
            line-height: 1;
            white-space: nowrap;
            font-size:.8em;
            background:#eee;
            }

            #footer ul
            {
            margin: 0;
            padding: 10px 0;
            padding-left: 0;
            overflow: hidden;
            color: #666;
            font-size:.8em;
            }

            #footer li
            {
            display: inline;
            padding: 0 4px;
            }

            .commonWidthForPermission
            {
            width: 1800px;
            }

            .table
            {
            width: auto;
            border: 1px solid #808080;
            }

            .tableRow
            {
            width: auto;
            height: 20px;
            border-bottom: 1px solid #808080;
            padding: 0px;
            margin: 0px;
            }

            .centerText
            {
            text-align: center;
            vertical-align: middle;
            }

            .centerVertical
            {
            vertical-align: middle;
            }

            .cellWidthSPBasePermissionValue
            {
            width: 250px;
            }

            .cellWidthSPBasePermissionDescription
            {
            width: 1200px;
            }

            .cellWidthSPBasePermissionName
            {
            width: 200px;
            }

            .cellHeight
            {
            height: 20px;
            }

          </style>
        </head>
        <body>
          <div class="wide">
            <xsl:apply-templates />
          </div>
          
          <div id="footer" class="clear round wide">
            <ul class="footer-nav">
              <li>Copyright Henri d'Orgeval 2011</li>
              <li>Beta Version {currentScriptVersion}</li>
              <li>
                <a href="http://SPPoster.codeplex.com" target="_codeplex">Get latest release on Codeplex</a>
              </li>
            </ul>
          </div>
        </body>
      </html>
  </xsl:template>

  <xsl:template match="Objs/Obj/MS">
      <fieldset class="round darkGray">
        <legend >
          SharePoint Farm Poster generated on {currentDate}
        </legend>
        <div class="left common2Width">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Global Farm Properties
            </legend>
            <div class="spFarm round indent">
              <xsl:apply-templates select="Obj[@N='Status']" />
              <xsl:apply-templates select="S[@N='ConfigurationDatabaseConnectionString']" />
              <xsl:apply-templates select="S[@N='ConfigurationDatabaseID']" />
              <xsl:apply-templates select="S[@N='CentralAdminUrl']" />
              <xsl:apply-templates select="S[@N='BuildVersion']" />
              <xsl:apply-templates select="S[@N='DisplayName']" />
              <xsl:apply-templates select="S[@N='Version']" />
              <xsl:apply-templates select="S[@N='LicenceType']" />
              <xsl:apply-templates select="S[@N='SharePointRoot']" />
              <xsl:apply-templates select="S[@N='DefaultServiceAccount']" />
              <xsl:apply-templates select="I32[@N='DaysBeforePasswordExpirationToSendEmail']" />
              <xsl:apply-templates select="B[@N='BrowserCEIPEnabled']" />
              <xsl:apply-templates select="B[@N='CEIPEnabled']" />
              
              <xsl:apply-templates select="Obj[@N='Properties']" />
              <xsl:apply-templates select="Obj[@N='ServerLanguage']" />
              <xsl:apply-templates select="Obj[@N='InstalledLanguages']" />
            </div>
          </fieldset>
        </div>


        <div class="left commonWidth">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Web Applications
            </legend>
            <div class="spFarm round indent">
              <xsl:apply-templates select="Obj[@N='WebApplicationsList']" />
            </div>
          </fieldset>
        </div>
        
        <div class="left commonWidth">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Developer Dashboard Settings
            </legend>
            <div class="spFarm round indent">
              <xsl:apply-templates select="Obj[@N='DeveloperDashBoardSettings']" />
            </div>
          </fieldset>
        </div>

        <div class="left commonWidth">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Log Settings
            </legend>
            <div class="spFarm round indent">
              <xsl:apply-templates select="Obj[@N='LogSettings']" />
            </div>
          </fieldset>
        </div>
        
        <hr class="clear"/>
        
        <div class="left wide">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Quota Templates (<xsl:value-of select="I32[@N='QuotaTemplatesCount']"/>)
            </legend>
            <xsl:apply-templates select="Obj[@N='QuotaTemplates']" />
          </fieldset>
        </div>

        <div class="left wide">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Deployed Farm Solutions (<xsl:value-of select="I32[@N='FarmSolutionsCount']"/>)
            </legend>
            <xsl:apply-templates select="Obj[@N='FarmSolutions']" />
          </fieldset>
        </div>
        
        <div class="left wide">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Farm-scoped Features (<xsl:value-of select="I32[@N='FarmFeaturesCount']"/>)
            </legend>
            <!--<div class="spFarmFeatures round indent">-->
              <xsl:apply-templates select="Obj[@N='FarmFeatures']" />
            <!--</div>-->
          </fieldset>
        </div>

        <div class="left wide">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Managed Accounts (<xsl:value-of select="I32[@N='ManagedAccountsCount']"/>)
            </legend>
            <!--<div class="spFarm round indent">-->
              <xsl:apply-templates select="Obj[@N='ManagedAccounts']" />
            <!--</div>-->
          </fieldset>
        </div>

        <div class="left wide">
          <fieldset class="round darkGray indent bgColorLightGreen">
            <legend>
              Service Application Proxy Groups (<xsl:value-of select="I32[@N='ServiceApplicationProxyGroupsCount']"/>)
            </legend>
            <!--<div class="spFarm round indent">-->
            <xsl:apply-templates select="Obj[@N='ServiceApplicationProxyGroups']" />
            <!--</div>-->
          </fieldset>
        </div>
       
        <div class="left wide">
          <fieldset class="round darkGray indent">
            <legend>
              Web Applications (<xsl:value-of select="I32[@N='WebApplicationsCount']"/>)
            </legend>
            <!--<div class="spFarm round indent">-->
              <xsl:apply-templates select="Obj[@N='WebApplications']" />
            <!--</div>-->
          </fieldset>
        </div>

        <div class="left wide">
          <fieldset class="round darkGray indent">
            <legend>
              All Site Collections (<xsl:value-of select="I32[@N='AllSiteCollectionsCount']"/>)
            </legend>
            <!--<div class="spFarm round indent">-->
            <xsl:apply-templates select="Obj[@N='AllSiteCollections']" />
            <!--</div>-->
          </fieldset>
        </div>
        
        
        
      </fieldset>
    
  </xsl:template>

  <xsl:template match="Obj[@N='AllSiteCollections']">
    <xsl:for-each select="LST/Obj/MS">
      <xsl:variable  name="siteCollectionUrl" select="S[@N='Url']"></xsl:variable>
      <xsl:variable  name="siteCollectionID" select="G[@N='ID']"></xsl:variable>
      <a>
        <xsl:attribute name='name'><xsl:value-of select="$siteCollectionID"/>-Level2</xsl:attribute>
      </a>
      <fieldset class="round darkGray indent bgColorLightGreen">
        <legend>
          <span class="property">
            <xsl:value-of select="$siteCollectionUrl"/>
          </span>
          (
          <a class="smallerFont">
            <xsl:attribute name='href'>#<xsl:value-of select="$siteCollectionID"/>-PermissionLevels</xsl:attribute>
            Permission Levels
          </a>
          )
        </legend>
       
        <!--<div class="left commonWidth">
          <fieldset class="round darkGray indent backgroundGreen">
            <legend>Blocked File Types</legend>
            <div class="round indent smallerFont">
              <xsl:apply-templates select="Obj[@N='BlockedFileExtensionsSettings']" />
            </div>
          </fieldset>
        </div>-->

        <div class="clear">
          <fieldset class="round darkGray indent">
            <legend>
              SiteCollection-scoped Features (<xsl:value-of select="I32[@N='SiteCollectionFeaturesCount']"/>)
            </legend>
            <xsl:apply-templates select="Obj[@N='SiteCollectionFeatures']" />
          </fieldset>
        </div>

        <div class="clear">
          <a>
            <xsl:attribute name='name'><xsl:value-of select="$siteCollectionID"/>-PermissionLevels</xsl:attribute>
          </a>
          <fieldset class="round darkGray indent">
            <legend>
              Permission Levels (<xsl:value-of select="I32[@N='PermissionLevelsCount']"/>)
            </legend>
            <xsl:apply-templates select="Obj[@N='PermissionLevels']" />
          </fieldset>
        </div>


        <div class="clear">
          <a>
            <xsl:attribute name='name'><xsl:value-of select="$siteCollectionID"/>-PermissionLevelsMatrix</xsl:attribute>
          </a>
          <fieldset class="round darkGray indent">
            <legend>
              Permission Levels Matrix 
            </legend>
            <xsl:apply-templates select="Obj[@N='RoleDefinitions']" />
            <xsl:apply-templates select="Obj[@N='PermissionLevelsMatrix']" />
          </fieldset>
        </div>

        <!--</div>-->
      </fieldset>
      <hr class="clear" />
    </xsl:for-each>
  </xsl:template>



  <xsl:template match="Obj[@N='RoleDefinitions']">
    <div class="table clear indent">
      <div class="tableRow clear backgroundGreen">
        
        <xsl:for-each select="LST/Obj/MS">
          <div class="left property cellWidthSPBasePermissionName centerText cellHeight smallerFont">
            <xsl:value-of select="S[@N='Name']"/>
          </div>
        </xsl:for-each>
        
        <div class="left property cellWidthSPBasePermissionValue centerText cellHeight smallerFont">SPBasePermissions Value</div>
        <div class="left property cellWidthSPBasePermissionDescription centerVertical cellHeight smallerFont">Description</div>
      </div>
    </div>
  </xsl:template>
  

  <xsl:template match="Obj[@N='PermissionLevelsMatrix']">
    <div class="table clear indent">
      
      <xsl:for-each select="LST/Obj/MS">
        
        <div class="tableRow clear">
          
          <xsl:for-each select="Obj/LST/Obj/MS">
            <xsl:variable  name="enabled" select="B[@N='Enabled']"></xsl:variable>
            <xsl:variable  name="permissionLevelName" select="S[@N='Name']"></xsl:variable>
            <div class="left cellWidthSPBasePermissionName smallerFont centerText cellHeight backgroundGreen">
              <xsl:if test="$enabled='false'">
                <xsl:attribute name='class'>left cellWidthSPBasePermissionName smallerFont centerText cellHeight backgroundRed</xsl:attribute>
              </xsl:if>
              <xsl:value-of select="$enabled"/>
            </div>
          </xsl:for-each>
          
          <div class="left cellWidthSPBasePermissionValue smallerFont centerText cellHeight">
            <xsl:value-of select="S[@N='Name']"/>
          </div>
          <div class="left cellWidthSPBasePermissionDescription smallerFont centerVertical cellHeight">
            <xsl:value-of select="S[@N='Description']"/>
          </div>
        </div>
      </xsl:for-each>

    </div>
    
  </xsl:template>
  
  
  <xsl:template match="Obj[@N='PermissionLevels']">
    <xsl:for-each select="LST/Obj/MS">
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen permissionLevelHeight">
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
            </span>
          </legend>
          <div class="round indent smallerFont">
            <xsl:apply-templates select="S[@N='Type']" />
            <xsl:apply-templates select="S[@N='Description']" />
          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template match="Obj[@N='WebApplicationsList']">
    <xsl:for-each select="LST/Obj/MS">
      <xsl:variable  name="webApplicationName" select="S[@N='Name']"></xsl:variable>
      <span class="property">
        <xsl:value-of select="$webApplicationName"/>
      </span>
      (
      <a class="smallerFont">
        <xsl:attribute name='href'>
          #<xsl:value-of select="$webApplicationName"/>
        </xsl:attribute>
        view details
      </a>
      )
      <hr/>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="S|I32|B|MS/B|MS/S|MS/I32|MS/U32">
    <span class="property">
      <xsl:value-of select="@N"/>
    </span>:
    <xsl:value-of select="."/>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='ServiceApplicationProxyGroups']">
    <xsl:for-each select="LST/Obj/MS">
      <div class="left wide">
        <fieldset class="round darkGray indent bgColorLightGreen">
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
            </span>
          </legend>
          <div class="round indent">
            <fieldset class="round darkGray indent bgColorLightGreen">
              <legend>
                Service Application Proxies that are included in this group (<xsl:value-of select="I32[@N='ServiceApplicationProxiesCount']"/>)
              </legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='ServiceApplicationProxies']" />
              </div>
            </fieldset>

          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='ServiceApplicationProxies']">
    <xsl:for-each select="LST/Obj/MS">
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen proxyHeight">
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
              <xsl:apply-templates select="S[@N='ManageProxyUrl']" />
            </span>
          </legend>
          <div class="round indent smallerFont">
            <xsl:apply-templates select="S[@N='Type']" />
            <xsl:apply-templates select="S[@N='Status']" />
            <xsl:apply-templates select="G[@N='Id']" />
            <xsl:apply-templates select="Obj[@N='ServiceApplication']" />
            
          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='ServiceApplication']">
    <fieldset class="round darkGray indent backgroundGreen">
      <legend>
        <span class="property">
          Associated Service Application
          <xsl:apply-templates select="MS/S[@N='ManageServiceApplicationUrl']" />
        </span>
      </legend>
      <div class="round indent smallerFont">
        <xsl:apply-templates select="MS/S[@N='Type']" />
        <xsl:apply-templates select="MS/S[@N='Status']" />
      </div>
    </fieldset>
  </xsl:template>

  <xsl:template match="MS/S[@N='ManageServiceApplicationUrl']">
    (<a target="CentralAdmin" class="smallerFont">
      <xsl:attribute name='href'>
        <xsl:value-of select="." />
      </xsl:attribute>
      Manage Service Application configuration
    </a>)
  </xsl:template>
  
  <xsl:template match="S[@N='ManageProxyUrl']">
    (<a target="CentralAdmin" class="smallerFont"><xsl:attribute name='href'><xsl:value-of select="." /></xsl:attribute>
      Manage Proxy configuration
    </a>)
  </xsl:template>
  
  <xsl:template match="Obj[@N='FarmSolutions']">
    <xsl:for-each select="LST/Obj/MS">
      <xsl:variable  name="solutionIsDeployed" select="B[@N='Deployed']"></xsl:variable>
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen farmSolutionHeight">
          <xsl:if test="$solutionIsDeployed='false'">
            <xsl:attribute name='class'>round darkGray indent backgroundRed farmSolutionHeight</xsl:attribute>
          </xsl:if>
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
            </span>
          </legend>
          <div class="round indent smallerFont">
            <xsl:apply-templates select="S" />
            <xsl:apply-templates select="B" />
            <!--<xsl:apply-templates select="S[@N='PasswordExpirationDate']" />
            <xsl:apply-templates select="S[@N='PasswordChangeSchedule']" />-->

          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='QuotaTemplates']">
    <xsl:for-each select="LST/Obj/MS">
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen quotaTemplateHeight">
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
            </span>
          </legend>
          <div class="round indent smallerFont">
            <xsl:apply-templates select="S" />
            <!--<xsl:apply-templates select="S[@N='PasswordExpirationDate']" />
            <xsl:apply-templates select="S[@N='PasswordChangeSchedule']" />-->

          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='WebApplications']">
    <xsl:for-each select="LST/Obj/MS">
      <xsl:variable  name="webApplicationName" select="S[@N='Name']"></xsl:variable>
      <a>
        <xsl:attribute name='name'>
          <xsl:value-of select="$webApplicationName"/>
        </xsl:attribute>
      </a>
        <fieldset class="round darkGray indent bgColorLightGreen">
          <legend>
            <span class="property">
              <xsl:value-of select="$webApplicationName"/>
              <!--<xsl:value-of select="S[@N='Name']"/>-->
            </span>
          </legend>
          <!--<div class="spFarm round indent smallerFont">-->
          <!--<xsl:apply-templates select="B[@N='EnableSharePointToAutomaticallyGenerateAndUpdatePassword']" />
              <xsl:apply-templates select="S[@N='PasswordLastChange']" />
              <xsl:apply-templates select="S[@N='PasswordExpirationDate']" />
              <xsl:apply-templates select="S[@N='PasswordChangeSchedule']" />-->
          
          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Global Settings</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='GlobalSettings']" />
              </div>
            </fieldset>
          </div>
          
          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Managed Paths</legend>
              <!--<div class="spFarm round indent smallerFont">-->
              <xsl:apply-templates select="Obj[@N='ManagedPaths']" />
              <!--</div>-->
            </fieldset>
          </div>

          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Large List Throttling Settings</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='LargeListThrottlingSettings']" />
              </div>
            </fieldset>
          </div>

          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>IIS Settings</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='IISSettings']" />
              </div>
            </fieldset>
          </div>

          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>SharePoint Designer Settings</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='SharePointDesignerSettings']" />
              </div>
            </fieldset>
          </div>

          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Web Part Settings</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='WebPartSettings']" />
              </div>
            </fieldset>
          </div>

          <!--<div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Inline MIME Type Settings</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='InlineDownloadedMimeTypesSettings']" />
              </div>
            </fieldset>
          </div>-->

          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Blocked File Types</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='BlockedFileExtensionsSettings']" />
              </div>
            </fieldset>
          </div>

          <div class="left commonWidthForPermission">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>Available Permissions for all Site Collections in this Web Application</legend>
              <div class="round indent smallerFont">
                <xsl:apply-templates select="Obj[@N='PermissionsSettings']" />
              </div>
            </fieldset>
          </div>

          <div class="left commonWidth">
            <fieldset class="round darkGray indent backgroundGreen">
              <legend>
                Security Policies (<xsl:value-of select="I32[@N='SecurityPoliciesCount']"/>)
              </legend>
              <!--<div class="spFarm round indent smallerFont">-->
              <xsl:apply-templates select="Obj[@N='SecurityPolicies']" />
              <!--</div>-->
            </fieldset>
          </div>

          <div class="clear">
            <fieldset class="round darkGray indent">
              <legend>
                Content Databases (<xsl:value-of select="I32[@N='ContentDatabasesCount']"/>)
              </legend>
              <xsl:apply-templates select="Obj[@N='ContentDatabaseSettings']" />
              <xsl:apply-templates select="Obj[@N='ContentDatabases']" />
            </fieldset>
          </div>
          
          <div class="clear">
            <fieldset class="round darkGray indent">
              <legend>
                Site Collections (<xsl:value-of select="I32[@N='WebApplicationSiteCollectionsCount']"/>)
              </legend>
              <xsl:apply-templates select="Obj[@N='WebApplicationSiteCollections']" />
            </fieldset>
          </div>

          <div class="clear">
            <fieldset class="round darkGray indent">
              <legend>
                WebApplication-scoped Features (<xsl:value-of select="I32[@N='WebApplicationFeaturesCount']"/>)
              </legend>
              <xsl:apply-templates select="Obj[@N='WebApplicationFeatures']" />
            </fieldset>
          </div>
          
          <!--</div>-->
        </fieldset>
        <hr class="clear" />
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='SecurityPolicies']">
    <xsl:for-each select="LST/Obj/MS">
      <fieldset class="round darkGray indent backgroundGreen">
        <legend>
          <span class="property">
            <xsl:value-of select="S[@N='DisplayName']"/>
          </span>
        </legend>
        <div class="round indent smallerFont">
          <xsl:apply-templates select="S[@N='Name']" />
          <xsl:apply-templates select="S[@N='Authorizations']" />
          <xsl:apply-templates select="B[@N='IsSiteCollectionAdministrator']" />
          <xsl:apply-templates select="B[@N='IsSiteCollectionAuditor']" />
        </div>
      </fieldset>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='PermissionsSettings']">
    <div class="table clear">
      <div class="tableRow clear">
        <div class="left property cellWidthSPBasePermissionName centerText cellHeight">RightsMask</div>
        <div class="left property cellWidthSPBasePermissionValue centerText cellHeight">SPBasePermissions Value</div>
        <div class="left property cellWidthSPBasePermissionDescription centerVertical cellHeight">Description</div>
      </div>

      <xsl:for-each select="MS/Obj[@N='RightsMask']/LST/Obj/MS">
        <xsl:variable  name="enabled" select="B[@N='Enabled']"></xsl:variable>
        <div class="tableRow clear">
          <xsl:if test="$enabled='false'">
            <xsl:attribute name='class'>tableRow clear backgroundRed</xsl:attribute>
          </xsl:if>
          <div class="left cellWidthSPBasePermissionName smallerFont centerText cellHeight">
            <xsl:value-of select="$enabled"/>
          </div>
          <div class="left cellWidthSPBasePermissionValue smallerFont centerText cellHeight">
            <xsl:value-of select="S[@N='Name']"/>
          </div>
          <div class="left cellWidthSPBasePermissionDescription smallerFont centerVertical cellHeight">
            <xsl:value-of select="S[@N='Description']"/>
          </div>
        </div>
      </xsl:for-each>
      
    </div>
    <!--<xsl:for-each select="LST/Obj/MS">
      <xsl:variable  name="featureIsActivated" select="B[@N='IsActivated']"></xsl:variable>
      <xsl:variable  name="featureExistsOnDisk" select="B[@N='ExistsOnDisk']"></xsl:variable>
      <xsl:variable  name="featureIsHidden" select="B[@N='Hidden']"></xsl:variable>
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen featureHeight">
          <xsl:if test="$featureIsActivated='false'">
            <xsl:attribute name='class'>round darkGray indent backgroundRed featureHeight</xsl:attribute>
          </xsl:if>
          <xsl:if test="$featureIsHidden='true' and $featureIsActivated='false'">
            <xsl:attribute name='class'>round darkGray indent backgroundGrayRed featureHeight</xsl:attribute>
          </xsl:if>
          <xsl:if test="$featureIsHidden='true' and $featureIsActivated='true'">
            <xsl:attribute name='class'>round darkGray indent backgroundGrayGreen featureHeight</xsl:attribute>
          </xsl:if>
          <xsl:if test="$featureExistsOnDisk='false'">
            <xsl:attribute name='class'>round darkGray indent backgroundRed featureHeight</xsl:attribute>
          </xsl:if>
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
            </span>
          </legend>
          <div class="round indent smallerFont">
            <span class="property">Title</span>:
            <xsl:value-of select="S[@N='Title']"/>
            <hr/>
            <xsl:apply-templates select="S[@N='Description']" />
            <xsl:apply-templates select="B[@N='Hidden']" />
            <xsl:apply-templates select="B[@N='IsActivated']" />
            <xsl:apply-templates select="B[@N='RequireResources']" />
            <xsl:apply-templates select="S[@N='DefaultResourceFile']" />
            <xsl:apply-templates select="S[@N='RootDirectory']" />
            <xsl:apply-templates select="G[@N='GUID']" />
            <xsl:apply-templates select="S[@N='Version']" />
          </div>
        </fieldset>
      </div>
    </xsl:for-each>-->
  </xsl:template>

  <xsl:template match="Obj[@N='WebPartSettings']">
    <xsl:apply-templates select="MS/B" />
    <xsl:apply-templates select="MS/S" />
    <xsl:apply-templates select="MS/I32" />
    <xsl:apply-templates select="MS/U32" />
    <fieldset class="round darkGray indent backgroundGreen">
      <legend>Scriptable Settings</legend>
      <div class="round indent smallerFont">
        <xsl:apply-templates select="MS/Obj[@N='ScriptableSettings']" />
      </div>
    </fieldset>
    
  </xsl:template>

  <xsl:template match="Obj[@N='ContentDatabases']">
    <xsl:for-each select="LST/Obj/MS">
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen smallerFont contentDBHeight">
          <legend>
            <span class="property">
              <xsl:value-of select="S[@N='Name']"/>
            </span>
          </legend>
          <div class="round indent smallerFont">
            <xsl:apply-templates select="S[@N='DatabaseServer']" />
            <xsl:apply-templates select="S[@N='ConnectionString']" />
            <xsl:apply-templates select="S[@N='CurrentNumberOfSiteCollection']" />
            <xsl:apply-templates select="S[@N='DiskUsageForAllSiteCollections']" />
            <xsl:apply-templates select="Obj[@N='SiteCollections']" />
          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template match="Obj[@N='ContentDatabaseSettings']">
    <div class="left commonWidth">
      <fieldset class="round darkGray indent backgroundGreen smallerFont contentDBHeight">
        <legend>
          <span class="property">
            <xsl:value-of select="MS/S[@N='Name']"/>
          </span>
        </legend>
        <div class="round indent smallerFont">
          <xsl:apply-templates select="MS/S[@N='DatabaseServer']" />
          <xsl:apply-templates select="MS/S[@N='ConnectionString']" />
          <xsl:apply-templates select="MS/S[@N='CurrentNumberOfSiteCollection']" />
          <xsl:apply-templates select="MS/S[@N='DiskUsageForAllSiteCollections']" />
          <xsl:apply-templates select="MS/Obj[@N='SiteCollections']" />
        </div>
      </fieldset>
    </div>
  </xsl:template>

  <xsl:template match="MS/Obj[@N='SiteCollections'] | Obj[@N='SiteCollections']">
    <span class="property">
      <xsl:value-of select="@N"/>
    </span>:
    <xsl:for-each select="LST/S">
      <xsl:if test="(position( )) > 1">
        ,
      </xsl:if>
      <xsl:value-of select="."/>
    </xsl:for-each>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='WebApplicationSiteCollections']">
    <xsl:for-each select="LST/Obj/MS">
      <xsl:variable  name="writeLocked" select="B[@N='UnavailableForWriteAccess']"></xsl:variable>
      <xsl:variable  name="readLocked" select="B[@N='UnavailableForReadAccess']"></xsl:variable>
      <xsl:variable  name="url" select="S[@N='Url']"></xsl:variable>
      <xsl:variable  name="siteCollectionID" select="G[@N='ID']"></xsl:variable>
      
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen siteCollectionHeight smallerFont">
          <xsl:if test="$readLocked='true'">
            <xsl:attribute name='class'>round darkGray indent backgroundRed siteCollectionHeight smallerFont</xsl:attribute>
          </xsl:if>
          <xsl:if test="$writeLocked='true'">
            <xsl:attribute name='class'>round darkGray indent backgroundOrange siteCollectionHeight smallerFont</xsl:attribute>
          </xsl:if>
          <legend>
            <span class="property">
              <a target="SPSiteWindow">
                <xsl:attribute name='href'><xsl:value-of select="$url"/></xsl:attribute>
                <xsl:value-of select="$url"/>
              </a></span>

            (
            <a class="smallerFont">
              <xsl:attribute name='href'>#<xsl:value-of select="$siteCollectionID"/>-Level2</xsl:attribute>
              more details
            </a>
            )
          </legend>
          <div class="round indent smallerFont">
            <xsl:apply-templates select="Nil[@N='ReadOnly']" />
            <xsl:apply-templates select="S[@N='CreationDate']" />
            <xsl:apply-templates select="S[@N='ContentDatabase']" />
            <xsl:apply-templates select="S[@N='Owner']" />
            <span class="property">Locked</span>: <xsl:value-of select="B[@N='UnavailableForReadAccess']"/>; 
            <span class="property">ReadOnly</span>: <xsl:value-of select="B[@N='ReadOnly']"/>;
            <span class="property">UnavailableForWriteAccess (can Read/Delete/Update)</span>: <xsl:value-of select="B[@N='UnavailableForWriteAccess']"/>
            <hr/>
            <xsl:apply-templates select="S[@N='RootWeb']" />
            <xsl:apply-templates select="S[@N='Language']" />
            <xsl:apply-templates select="B[@N='AllowUnsafeUpdates']" />
            <xsl:apply-templates select="B[@N='LetSharePointTrapAndHandleAccessDeniedExceptions']" />
            <xsl:apply-templates select="I32" />
            <xsl:apply-templates select="S[@N='Quota']" />
            <xsl:apply-templates select="S[@N='DiskStorageUsed']" />
            <xsl:apply-templates select="S[@N='ComputerResourceUsed']" />
            <xsl:apply-templates select="S[@N='UserSolutions']" />
            <xsl:apply-templates select="S[@N='Users']" />
            <xsl:apply-templates select="S[@N='UsersActivityForTheCurrentMonth']" />
            <xsl:apply-templates select="S[@N='WebsAndListsWithUniquePermission']" />
            
          </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='IISSettings']">
    <xsl:apply-templates select="MS/B" />
    <xsl:apply-templates select="MS/S" />
    <xsl:apply-templates select="MS/I32" />
    <xsl:apply-templates select="MS/U32" />
    <fieldset class="round darkGray indent backgroundGreen">
      <legend>Authentication Settings</legend>
      <div class="round indent smallerFont">
        <xsl:apply-templates select="MS/Obj[@N='AuthenticationSettings']" />
      </div>
    </fieldset>
  </xsl:template>

  <xsl:template match="Nil[@N='ReadOnly']">
    <span class="property backgroundRed">
      This Site Collection seems not work properly
    </span>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='ManagedPaths']">
    <xsl:for-each select="LST/Obj/MS">
      <fieldset class="round darkGray indent backgroundGreen">
        <legend>
          <span class="property"><xsl:value-of select="S[@N='Path']"/></span>
        </legend>
        <div class="round indent smallerFont">
          <xsl:apply-templates select="S[@N='PathType']" />
          <xsl:apply-templates select="S[@N='Description']" />
        </div>
      </fieldset>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='ManagedAccounts']">
    <xsl:for-each select="LST/Obj/MS">
      <div class="left commonWidth">
        <fieldset class="round darkGray indent backgroundGreen">
          <legend>
            <span class="property"><xsl:value-of select="S[@N='UserName']"/></span>
          </legend>
          <div class="round indent smallerFont">
          <xsl:apply-templates select="B[@N='EnableSharePointToAutomaticallyGenerateAndUpdatePassword']" />
          <xsl:apply-templates select="S[@N='PasswordLastChange']" />
          <xsl:apply-templates select="S[@N='PasswordExpirationDate']" />
          <xsl:apply-templates select="S[@N='PasswordChangeSchedule']" />
          
          <fieldset class="round darkGray indent backgroundGreen">
            <legend>SharePoint Services that run under this Identity</legend>
            <div class="round indent smallerFont">
              <xsl:apply-templates select="Obj[@N='Services']" />
            </div>
          </fieldset>
          
        </div>
        </fieldset>
      </div>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="Obj[@N='Services']">
    <xsl:for-each select="LST/Obj/MS">
      <fieldset class="round darkGray indent backgroundGreen">
        <legend>
          <span class="property"><xsl:value-of select="S[@N='Name']"/></span>
        </legend>
        <div class="round indent smallerFont">
          <xsl:apply-templates select="S[@N='Description']" />
          <xsl:apply-templates select="S[@N='Type']" />
        </div>
      </fieldset>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template match="Obj[@N='FarmFeatures'] | Obj[@N='WebApplicationFeatures'] | Obj[@N='SiteCollectionFeatures']">
      <xsl:for-each select="LST/Obj/MS">
        <xsl:variable  name="featureIsActivated" select="B[@N='IsActivated']"></xsl:variable>
        <xsl:variable  name="featureExistsOnDisk" select="B[@N='ExistsOnDisk']"></xsl:variable>
        <xsl:variable  name="featureIsHidden" select="B[@N='Hidden']"></xsl:variable>
        <xsl:variable  name="featureVersion" select="S[@N='Version']"></xsl:variable>
        <xsl:variable  name="featureVersionOnDisk" select="S[@N='VersionOnDisk']"></xsl:variable>
        <div class="left commonWidth">
          <fieldset class="round darkGray indent backgroundGreen featureHeight">
          <xsl:if test="$featureIsActivated='false'">
            <xsl:attribute name='class'>round darkGray indent backgroundRed featureHeight</xsl:attribute>
          </xsl:if>
            <xsl:if test="$featureIsHidden='true' and $featureIsActivated='false'">
              <xsl:attribute name='class'>round darkGray indent backgroundGrayRed featureHeight</xsl:attribute>
            </xsl:if>
            <xsl:if test="$featureIsHidden='true' and $featureIsActivated='true'">
              <xsl:attribute name='class'>round darkGray indent backgroundGrayGreen featureHeight</xsl:attribute>
            </xsl:if>
            <xsl:if test="$featureExistsOnDisk='false'">
              <xsl:attribute name='class'>round darkGray indent backgroundRed featureHeight</xsl:attribute>
            </xsl:if>

            <xsl:if test="$featureVersion != $featureVersionOnDisk">
              <xsl:attribute name='class'>round darkGray indent backgroundYellow featureHeight</xsl:attribute>
            </xsl:if>
            
          <legend>
            <span class="property"><xsl:value-of select="S[@N='Name']"/></span>
          </legend>
          <div class="round indent smallerFont">
            <span class="property">Title</span>:
            <xsl:value-of select="S[@N='Title']"/>
            <hr/>
            <xsl:apply-templates select="S[@N='Description']" />
            <xsl:apply-templates select="B[@N='Hidden']" />
            <xsl:apply-templates select="B[@N='IsActivated']" />
            <xsl:apply-templates select="B[@N='RequireResources']" />
            <xsl:apply-templates select="S[@N='DefaultResourceFile']" />
            <xsl:apply-templates select="S[@N='RootDirectory']" />
            <xsl:apply-templates select="G[@N='GUID']" />
            <xsl:apply-templates select="S[@N='Version']" />
            <xsl:if test="$featureVersion != $featureVersionOnDisk">
              <xsl:apply-templates select="S[@N='VersionOnDisk']" />
            </xsl:if>
          </div>
        </fieldset>
        </div>
      </xsl:for-each>
  </xsl:template>

  <xsl:template match="S[@N='Description']|S[@N='RootDirectory']|S[@N='DefaultResourceFile']|G[@N='GUID']|S[@N='Type']|S[@N='PasswordLastChange']|S[@N='PathType'] | S[@N='Version'] | S[@N='Owner'] | S[@N='DiskStorageUsed'] | S[@N='Quota'] | S[@N='ComputerResourceUsed'] | MS/S[@N='DatabaseServer'] | S[@N='DatabaseServer'] | MS/S[@N='CurrentNumberOfSiteCollection'] | S[@N='CurrentNumberOfSiteCollection'] | MS/S[@N='DiskUsageForAllSiteCollections'] | S[@N='DiskUsageForAllSiteCollections'] | S[@N='RootWeb'] | MS/S[@N='ConnectionString'] | S[@N='ConnectionString'] | S[@N='Type'] | S[@N='Status'] | MS/S[@N='Type'] | MS/S[@N='Status']">
    <span class="property"><xsl:value-of select="@N"/></span>:
    <xsl:value-of select="."/>
    <hr/>
  </xsl:template>

  <xsl:template match="B[@N='Hidden']|B[@N='IsActivated']|B[@N='ExistsOnDisk']|B[@N='RequireResources']|B[@N='EnableSharePointToAutomaticallyGenerateAndUpdatePassword'] | B[@N='AllowUnsafeUpdates'] | B[@N='LetSharePointTrapAndHandleAccessDeniedExceptions']">
    <span class="property"><xsl:value-of select="@N"/></span>:
    <xsl:value-of select="."/>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='Properties']">
    <span class="property">Property Bag (Key=Value)</span>:
    <xsl:for-each select="MS/S">
      <xsl:if test="(position( )) > 1">
        , 
      </xsl:if>
      <span class="property"><xsl:value-of select="@N"/></span>=<xsl:value-of select="."/>
    </xsl:for-each>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='Status']">
    <span class="property"><xsl:value-of select="@N"/></span>:
    <xsl:value-of select="./ToString"/>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='ServerLanguage']">
    <span class="property"><xsl:value-of select="@N"/></span>:
    <xsl:value-of select="./Obj/MS/S/@N"/>
    <xsl:variable  name="lcid" select="./Obj/MS/S/@N"></xsl:variable>
    <xsl:call-template name="ShowLCID">
      <xsl:with-param name="lcid" select="$lcid" />
    </xsl:call-template>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='InstalledLanguages']">
    <span class="property"><xsl:value-of select="@N"/></span>:
    <xsl:for-each select="./Obj/MS/S">
      <xsl:if test="(position( )) > 1">
        ,
      </xsl:if>
      <xsl:value-of select="@N"/>
      <xsl:variable  name="lcid" select="@N" />
      <xsl:call-template name="ShowLCID">
        <xsl:with-param name="lcid" select="$lcid" />
      </xsl:call-template>
    </xsl:for-each>
    <hr/>
  </xsl:template>

  <xsl:template match="Obj[@N='DeveloperDashBoardSettings'] | Obj[@N='LogSettings']|Obj[@N='LargeListThrottlingSettings']| MS/Obj[@N='AuthenticationSettings']| Obj[@N='SharePointDesignerSettings'] | Obj[@N='GlobalSettings'] | Obj[@N='InlineDownloadedMimeTypesSettings'] | Obj[@N='BlockedFileExtensionsSettings'] | MS/Obj[@N='ScriptableSettings']">
    <xsl:apply-templates select="MS/B" />
    <xsl:apply-templates select="MS/S" />
    <xsl:apply-templates select="MS/I32" />
    <xsl:apply-templates select="MS/U32" />
  </xsl:template>

  <xsl:template name="ShowLCID">
    <xsl:param name="lcid"  />
    <xsl:choose>
      <xsl:when test='$lcid=1025'> [Arabic (Saudi Arabia)]</xsl:when>
      <xsl:when test='$lcid=1026'> [Bulgarian (Bulgaria)]</xsl:when>
      <xsl:when test='$lcid=1027'> [Catalan (Catalan)]</xsl:when>
      <xsl:when test='$lcid=1028'> [Chinese (Taiwan)]</xsl:when>
      <xsl:when test='$lcid=1029'> [Czech (Czech Republic)]</xsl:when>
      <xsl:when test='$lcid=1030'> [Danish (Denmark)]</xsl:when>
      <xsl:when test='$lcid=1031'> [German (Germany)]</xsl:when>
      <xsl:when test='$lcid=1032'> [Greek (Greece)]</xsl:when>
      <xsl:when test='$lcid=1033'> [English (United States)]</xsl:when>
      <xsl:when test='$lcid=1035'> [Finnish (Finland)]</xsl:when>
      <xsl:when test='$lcid=1036'> [French (France)]</xsl:when>
      <xsl:when test='$lcid=1037'> [Hebrew (Israel)]</xsl:when>
      <xsl:when test='$lcid=1038'> [Hungarian (Hungary)]</xsl:when>
      <xsl:when test='$lcid=1039'> [Icelandic (Iceland)]</xsl:when>
      <xsl:when test='$lcid=1040'> [Italian (Italy)]</xsl:when>
      <xsl:when test='$lcid=1041'> [Japanese (Japan)]</xsl:when>
      <xsl:when test='$lcid=1042'> [Korean (Korea)]</xsl:when>
      <xsl:when test='$lcid=1043'> [Dutch (Netherlands)]</xsl:when>
      <xsl:when test='$lcid=1044'> [Norwegian, Bokmål (Norway)]</xsl:when>
      <xsl:when test='$lcid=1045'> [Polish (Poland)]</xsl:when>
      <xsl:when test='$lcid=1046'> [Portuguese (Brazil)]</xsl:when>
      <xsl:when test='$lcid=1048'> [Romanian (Romania)]</xsl:when>
      <xsl:when test='$lcid=1049'> [Russian (Russia)]</xsl:when>
      <xsl:when test='$lcid=1050'> [Croatian (Croatia)]</xsl:when>
      <xsl:when test='$lcid=1051'> [Slovak (Slovakia)]</xsl:when>
      <xsl:when test='$lcid=1052'> [Albanian (Albania)]</xsl:when>
      <xsl:when test='$lcid=1053'> [Swedish (Sweden)]</xsl:when>
      <xsl:when test='$lcid=1054'> [Thai (Thailand)]</xsl:when>
      <xsl:when test='$lcid=1055'> [Turkish (Turkey)]</xsl:when>
      <xsl:when test='$lcid=1056'> [Urdu (Islamic Republic of Pakistan)]</xsl:when>
      <xsl:when test='$lcid=1057'> [Indonesian (Indonesia)]</xsl:when>
      <xsl:when test='$lcid=1058'> [Ukrainian (Ukraine)]</xsl:when>
      <xsl:when test='$lcid=1059'> [Belarusian (Belarus)]</xsl:when>
      <xsl:when test='$lcid=1060'> [Slovenian (Slovenia)]</xsl:when>
      <xsl:when test='$lcid=1061'> [Estonian (Estonia)]</xsl:when>
      <xsl:when test='$lcid=1062'> [Latvian (Latvia)]</xsl:when>
      <xsl:when test='$lcid=1063'> [Lithuanian (Lithuania)]</xsl:when>
      <xsl:when test='$lcid=1065'> [Persian (Iran)]</xsl:when>
      <xsl:when test='$lcid=1066'> [Vietnamese (Vietnam)]</xsl:when>
      <xsl:when test='$lcid=1067'> [Armenian (Armenia)]</xsl:when>
      <xsl:when test='$lcid=1068'> [Azeri (Latin, Azerbaijan)]</xsl:when>
      <xsl:when test='$lcid=1069'> [Basque (Basque)]</xsl:when>
      <xsl:when test='$lcid=1071'> [Macedonian (Former Yugoslav Republic of Macedonia)]</xsl:when>
      <xsl:when test='$lcid=1078'> [Afrikaans (South Africa)]</xsl:when>
      <xsl:when test='$lcid=1079'> [Georgian (Georgia)]</xsl:when>
      <xsl:when test='$lcid=1080'> [Faroese (Faroe Islands)]</xsl:when>
      <xsl:when test='$lcid=1081'> [Hindi (India)]</xsl:when>
      <xsl:when test='$lcid=1086'> [Malay (Malaysia)]</xsl:when>
      <xsl:when test='$lcid=1087'> [Kazakh (Kazakhstan)]</xsl:when>
      <xsl:when test='$lcid=1088'> [Kyrgyz (Kyrgyzstan)]</xsl:when>
      <xsl:when test='$lcid=1089'> [Kiswahili (Kenya)]</xsl:when>
      <xsl:when test='$lcid=1091'> [Uzbek (Latin, Uzbekistan)]</xsl:when>
      <xsl:when test='$lcid=1092'> [Tatar (Russia)]</xsl:when>
      <xsl:when test='$lcid=1094'> [Punjabi (India)]</xsl:when>
      <xsl:when test='$lcid=1095'> [Gujarati (India)]</xsl:when>
      <xsl:when test='$lcid=1097'> [Tamil (India)]</xsl:when>
      <xsl:when test='$lcid=1098'> [Telugu (India)]</xsl:when>
      <xsl:when test='$lcid=1099'> [Kannada (India)]</xsl:when>
      <xsl:when test='$lcid=1102'> [Marathi (India)]</xsl:when>
      <xsl:when test='$lcid=1103'> [Sanskrit (India)]</xsl:when>
      <xsl:when test='$lcid=1104'> [Mongolian (Cyrillic, Mongolia)]</xsl:when>
      <xsl:when test='$lcid=1110'> [Galician (Galician)]</xsl:when>
      <xsl:when test='$lcid=1111'> [Konkani (India)]</xsl:when>
      <xsl:when test='$lcid=1114'> [Syriac (Syria)]</xsl:when>
      <xsl:when test='$lcid=1125'> [Divehi (Maldives)]</xsl:when>
      <xsl:when test='$lcid=2049'> [Arabic (Iraq)]</xsl:when>
      <xsl:when test='$lcid=2052'> [Chinese (People's Republic of China)]</xsl:when>
      <xsl:when test='$lcid=2055'> [German (Switzerland)]</xsl:when>
      <xsl:when test='$lcid=2057'> [English (United Kingdom)]</xsl:when>
      <xsl:when test='$lcid=2058'> [Spanish (Mexico)]</xsl:when>
      <xsl:when test='$lcid=2060'> [French (Belgium)]</xsl:when>
      <xsl:when test='$lcid=2064'> [Italian (Switzerland)]</xsl:when>
      <xsl:when test='$lcid=2067'> [Dutch (Belgium)]</xsl:when>
      <xsl:when test='$lcid=2068'> [Norwegian, Nynorsk (Norway)]</xsl:when>
      <xsl:when test='$lcid=2070'> [Portuguese (Portugal)]</xsl:when>
      <xsl:when test='$lcid=2074'> [Serbian (Latin, Serbia and Montenegro (Former))]</xsl:when>
      <xsl:when test='$lcid=2077'> [Swedish (Finland)]</xsl:when>
      <xsl:when test='$lcid=2092'> [Azeri (Cyrillic, Azerbaijan)]</xsl:when>
      <xsl:when test='$lcid=2110'> [Malay (Brunei Darussalam)]</xsl:when>
      <xsl:when test='$lcid=2115'> [Uzbek (Cyrillic, Uzbekistan)]</xsl:when>
      <xsl:when test='$lcid=3073'> [Arabic (Egypt)]</xsl:when>
      <xsl:when test='$lcid=3076'> [Chinese (Hong Kong S.A.R.)]</xsl:when>
      <xsl:when test='$lcid=3079'> [German (Austria)]</xsl:when>
      <xsl:when test='$lcid=3081'> [English (Australia)]</xsl:when>
      <xsl:when test='$lcid=3082'> [Spanish (Spain)]</xsl:when>
      <xsl:when test='$lcid=3084'> [French (Canada)]</xsl:when>
      <xsl:when test='$lcid=3098'> [Serbian (Cyrillic, Serbia and Montenegro (Former))]</xsl:when>
      <xsl:when test='$lcid=4097'> [Arabic (Libya)]</xsl:when>
      <xsl:when test='$lcid=4100'> [Chinese (Singapore)]</xsl:when>
      <xsl:when test='$lcid=4103'> [German (Luxembourg)]</xsl:when>
      <xsl:when test='$lcid=4105'> [English (Canada)]</xsl:when>
      <xsl:when test='$lcid=4106'> [Spanish (Guatemala)]</xsl:when>
      <xsl:when test='$lcid=4108'> [French (Switzerland)]</xsl:when>
      <xsl:when test='$lcid=5121'> [Arabic (Algeria)]</xsl:when>
      <xsl:when test='$lcid=5124'> [Chinese (Macao S.A.R.)]</xsl:when>
      <xsl:when test='$lcid=5127'> [German (Liechtenstein)]</xsl:when>
      <xsl:when test='$lcid=5129'> [English (New Zealand)]</xsl:when>
      <xsl:when test='$lcid=5130'> [Spanish (Costa Rica)]</xsl:when>
      <xsl:when test='$lcid=5132'> [French (Luxembourg)]</xsl:when>
      <xsl:when test='$lcid=6145'> [Arabic (Morocco)]</xsl:when>
      <xsl:when test='$lcid=6153'> [English (Ireland)]</xsl:when>
      <xsl:when test='$lcid=6154'> [Spanish (Panama)]</xsl:when>
      <xsl:when test='$lcid=6156'> [French (Principality of Monaco)]</xsl:when>
      <xsl:when test='$lcid=7169'> [Arabic (Tunisia)]</xsl:when>
      <xsl:when test='$lcid=7177'> [English (South Africa)]</xsl:when>
      <xsl:when test='$lcid=7178'> [Spanish (Dominican Republic)]</xsl:when>
      <xsl:when test='$lcid=8193'> [Arabic (Oman)]</xsl:when>
      <xsl:when test='$lcid=8201'> [English (Jamaica)]</xsl:when>
      <xsl:when test='$lcid=8202'> [Spanish (Venezuela)]</xsl:when>
      <xsl:when test='$lcid=9217'> [Arabic (Yemen)]</xsl:when>
      <xsl:when test='$lcid=9225'> [English (Caribbean)]</xsl:when>
      <xsl:when test='$lcid=9226'> [Spanish (Colombia)]</xsl:when>
      <xsl:when test='$lcid=10241'> [Arabic (Syria)]</xsl:when>
      <xsl:when test='$lcid=10249'> [English (Belize)]</xsl:when>
      <xsl:when test='$lcid=10250'> [Spanish (Peru)]</xsl:when>
      <xsl:when test='$lcid=11265'> [Arabic (Jordan)]</xsl:when>
      <xsl:when test='$lcid=11273'> [English (Trinidad and Tobago)]</xsl:when>
      <xsl:when test='$lcid=11274'> [Spanish (Argentina)]</xsl:when>
      <xsl:when test='$lcid=12289'> [Arabic (Lebanon)]</xsl:when>
      <xsl:when test='$lcid=12297'> [English (Zimbabwe)]</xsl:when>
      <xsl:when test='$lcid=12298'> [Spanish (Ecuador)]</xsl:when>
      <xsl:when test='$lcid=13313'> [Arabic (Kuwait)]</xsl:when>
      <xsl:when test='$lcid=13321'> [English (Republic of the Philippines)]</xsl:when>
      <xsl:when test='$lcid=13322'> [Spanish (Chile)]</xsl:when>
      <xsl:when test='$lcid=14337'> [Arabic (U.A.E.)]</xsl:when>
      <xsl:when test='$lcid=14346'> [Spanish (Uruguay)]</xsl:when>
      <xsl:when test='$lcid=15361'> [Arabic (Bahrain)]</xsl:when>
      <xsl:when test='$lcid=15370'> [Spanish (Paraguay)]</xsl:when>
      <xsl:when test='$lcid=16385'> [Arabic (Qatar)]</xsl:when>
      <xsl:when test='$lcid=16394'> [Spanish (Bolivia)]</xsl:when>
      <xsl:when test='$lcid=17418'> [Spanish (El Salvador)]</xsl:when>
      <xsl:when test='$lcid=18442'> [Spanish (Honduras)]</xsl:when>
      <xsl:when test='$lcid=19466'> [Spanish (Nicaragua)]</xsl:when>
      <xsl:when test='$lcid=20490'> [Spanish (Puerto Rico)]</xsl:when>
      <xsl:when test='$lcid=31748'> [Chinese (Traditional)]</xsl:when>
      <xsl:when test='$lcid=31770'> [Serbian]</xsl:when>
      <xsl:when test='$lcid=1118'> [Amharic (Ethiopia)]</xsl:when>
      <xsl:when test='$lcid=2143'> [Tamazight (Latin) (Algeria)]</xsl:when>
      <xsl:when test='$lcid=2141'> [Inuktitut (Latin) (Canada)]</xsl:when>
      <xsl:when test='$lcid=6203'> [Sami (Southern) (Norway)]</xsl:when>
      <xsl:when test='$lcid=2128'> [Mongolian (Traditional Mongolian) (People's Republic of China)]</xsl:when>
      <xsl:when test='$lcid=1169'> [Scottish Gaelic (United Kingdom)]</xsl:when>
      <xsl:when test='$lcid=17417'> [English (Malaysia)]</xsl:when>
      <xsl:when test='$lcid=1164'> [Dari (Afghanistan)]</xsl:when>
      <xsl:when test='$lcid=2117'> [Bengali (Bangladesh)]</xsl:when>
      <xsl:when test='$lcid=1160'> [Wolof (Senegal)]</xsl:when>
      <xsl:when test='$lcid=1159'> [Kinyarwanda (Rwanda)]</xsl:when>
      <xsl:when test='$lcid=1158'> [K'iche (Guatemala)]</xsl:when>
      <xsl:when test='$lcid=1157'> [Yakut (Russia)]</xsl:when>
      <xsl:when test='$lcid=1156'> [Alsatian (France)]</xsl:when>
      <xsl:when test='$lcid=1155'> [Corsican (France)]</xsl:when>
      <xsl:when test='$lcid=1154'> [Occitan (France)]</xsl:when>
      <xsl:when test='$lcid=1153'> [Maori (New Zealand)]</xsl:when>
      <xsl:when test='$lcid=2108'> [Irish (Ireland)]</xsl:when>
      <xsl:when test='$lcid=2107'> [Sami (Northern) (Sweden)]</xsl:when>
      <xsl:when test='$lcid=1150'> [Breton (France)]</xsl:when>
      <xsl:when test='$lcid=9275'> [Sami (Inari) (Finland)]</xsl:when>
      <xsl:when test='$lcid=1148'> [Mohawk (Canada)]</xsl:when>
      <xsl:when test='$lcid=1146'> [Mapudungun (Chile)]</xsl:when>
      <xsl:when test='$lcid=1144'> [Yi (People's Republic of China)]</xsl:when>
      <xsl:when test='$lcid=2094'> [Lower Sorbian (Germany)]</xsl:when>
      <xsl:when test='$lcid=1136'> [Igbo (Nigeria)]</xsl:when>
      <xsl:when test='$lcid=1135'> [Greenlandic (Greenland)]</xsl:when>
      <xsl:when test='$lcid=1134'> [Luxembourgish (Luxembourg)]</xsl:when>
      <xsl:when test='$lcid=1133'> [Bashkir (Russia)]</xsl:when>
      <xsl:when test='$lcid=1132'> [Sesotho sa Leboa (South Africa)]</xsl:when>
      <xsl:when test='$lcid=1131'> [Quechua (Bolivia)]</xsl:when>
      <xsl:when test='$lcid=1130'> [Yoruba (Nigeria)]</xsl:when>
      <xsl:when test='$lcid=1128'> [Hausa (Latin) (Nigeria)]</xsl:when>
      <xsl:when test='$lcid=1124'> [Filipino (Philippines)]</xsl:when>
      <xsl:when test='$lcid=1123'> [Pashto (Afghanistan)]</xsl:when>
      <xsl:when test='$lcid=1122'> [Frisian (Netherlands)]</xsl:when>
      <xsl:when test='$lcid=1121'> [Nepali (Nepal)]</xsl:when>
      <xsl:when test='$lcid=1083'> [Sami (Northern) (Norway)]</xsl:when>
      <xsl:when test='$lcid=1117'> [Inuktitut (Syllabics) (Canada)]</xsl:when>
      <xsl:when test='$lcid=9242'> [Serbian (Latin) (Serbia)]</xsl:when>
      <xsl:when test='$lcid=1115'> [Sinhala (Sri Lanka)]</xsl:when>
      <xsl:when test='$lcid=10266'> [Serbian (Cyrillic) (Serbia)]</xsl:when>
      <xsl:when test='$lcid=1108'> [Lao (Lao P.D.R.)]</xsl:when>
      <xsl:when test='$lcid=1107'> [Khmer (Cambodia)]</xsl:when>
      <xsl:when test='$lcid=1106'> [Welsh (United Kingdom)]</xsl:when>
      <xsl:when test='$lcid=1105'> [Tibetan (People's Republic of China)]</xsl:when>
      <xsl:when test='$lcid=8251'> [Sami (Skolt) (Finland)]</xsl:when>
      <xsl:when test='$lcid=1101'> [Assamese (India)]</xsl:when>
      <xsl:when test='$lcid=1100'> [Malayalam (India)]</xsl:when>
      <xsl:when test='$lcid=16393'> [English (India)]</xsl:when>
      <xsl:when test='$lcid=1096'> [Oriya (India)]</xsl:when>
      <xsl:when test='$lcid=1093'> [Bengali (India)]</xsl:when>
      <xsl:when test='$lcid=1090'> [Turkmen (Turkmenistan)]</xsl:when>
      <xsl:when test='$lcid=5146'> [Bosnian (Latin) (Bosnia and Herzegovina)]</xsl:when>
      <xsl:when test='$lcid=1082'> [Maltese (Malta)]</xsl:when>
      <xsl:when test='$lcid=12314'> [Serbian (Cyrillic) (Montenegro)]</xsl:when>
      <xsl:when test='$lcid=3131'> [Sami (Northern) (Finland)]</xsl:when>
      <xsl:when test='$lcid=1077'> [isiZulu (South Africa)]</xsl:when>
      <xsl:when test='$lcid=1076'> [isiXhosa (South Africa)]</xsl:when>
      <xsl:when test='$lcid=1074'> [Setswana (South Africa)]</xsl:when>
      <xsl:when test='$lcid=1070'> [Upper Sorbian (Germany)]</xsl:when>
      <xsl:when test='$lcid=8218'> [Bosnian (Cyrillic) (Bosnia and Herzegovina)]</xsl:when>
      <xsl:when test='$lcid=1064'> [Tajik (Cyrillic) (Tajikistan)]</xsl:when>
      <xsl:when test='$lcid=6170'> [Serbian (Latin) (Bosnia and Herzegovina)]</xsl:when>
      <xsl:when test='$lcid=4155'> [Sami (Lule) (Norway)]</xsl:when>
      <xsl:when test='$lcid=1047'> [Romansh (Switzerland)]</xsl:when>
      <xsl:when test='$lcid=5179'> [Sami (Lule) (Sweden)]</xsl:when>
      <xsl:when test='$lcid=2155'> [Quechua (Ecuador)]</xsl:when>
      <xsl:when test='$lcid=3179'> [Quechua (Peru)]</xsl:when>
      <xsl:when test='$lcid=4122'> [Croatian (Latin) (Bosnia and Herzegovina)]</xsl:when>
      <xsl:when test='$lcid=11290'> [Serbian (Latin) (Montenegro)]</xsl:when>
      <xsl:when test='$lcid=7227'> [Sami (Southern) (Sweden)]</xsl:when>
      <xsl:when test='$lcid=18441'> [English (Singapore)]</xsl:when>
      <xsl:when test='$lcid=1152'> [Uyghur (People's Republic of China)]</xsl:when>
      <xsl:when test='$lcid=7194'> [Serbian (Cyrillic) (Bosnia and Herzegovina)]</xsl:when>
      <xsl:when test='$lcid=21514'> [Spanish (United States)]</xsl:when>

      <xsl:otherwise>
        (unknown LCID)
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="Objs/Obj/TN">
    
  </xsl:template>

</xsl:stylesheet>
'@

$xsl = $xsl -replace "{currentDate}", [System.DateTime]::Now.ToString([System.Globalization.CultureInfo]::CurrentUICulture)
$xsl = $xsl -replace "{currentScriptVersion}", "$ScriptVersion"

# Write the XSL to disk
$xslFilePath = "$Home\SPClonedFarm.xsl"
[System.IO.File]::WriteAllText($xslFilePath,$xsl)

WaitFor2Seconds
# End activity

###########################################################
# Begin activity
WriteProgress  "Preparing XSLT ..." 80
$xslt = New-Object System.Xml.Xsl.XslCompiledTransform
# Wait for the file system to be stable
[System.Threading.Thread]::Sleep(3000)
$xslt.Load($xslFilePath)
# End activity

###########################################################
# Begin activity
WriteProgress  "Generating HTML ..."  90
$htmlFilePath = "$Home\SPFarmPoster.html"
$xslt.Transform($xmlFilePath2,$htmlFilePath)

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress  "Loading HTML ..."  99
explorer $htmlFilePath

WaitFor2Seconds
# End activity


###########################################################
# Begin activity
WriteProgress  "HTML loaded" 100

WaitFor2Seconds
# End activity
