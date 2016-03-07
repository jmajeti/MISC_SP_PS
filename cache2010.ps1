# Clear the SharePoint Timer Cache
#
# 2009 Mickey Kamp Parbst Jervin (mickeyjervin.wordpress.com)
# 2011 Adapted by Nick Hobbs (nickhobbs.wordpress.com) to work with SharePoint 2010,
#      display more progress information, restart all timer services in the farm,
#      and make reusable functions.
 
# Output program information
Write-Host -foregroundcolor White ""
Write-Host -foregroundcolor White "Clear SharePoint Timer Cache"
 
#**************************************************************************************
# References
#**************************************************************************************
[void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint.Administration")
[void][reflection.assembly]::LoadWithPartialName("System")
[void][reflection.assembly]::LoadWithPartialName("System.Collections")
#**************************************************************************************
 
#**************************************************************************************
# Constants
#**************************************************************************************
Set-Variable timerServiceName -option Constant -value "SharePoint 2010 Timer"
Set-Variable timerServiceInstanceName -option Constant -value "Microsoft SharePoint Foundation Timer"
 
#**************************************************************************************
# Functions
#**************************************************************************************
 
#<summary>
# Stops the SharePoint Timer Service on each server in the SharePoint Farm.
#</summary>
#<param name="$farm">The SharePoint farm object.</param>
function StopSharePointTimerServicesInFarm([Microsoft.SharePoint.Administration.SPFarm]$farm)
{
    Write-Host ""
     
    # Iterate through each server in the farm, and each service in each server
    foreach($server in $farm.Servers)
    {
        foreach($instance in $server.ServiceInstances)
        {
            # If the server has the timer service then stop the service
            if($instance.TypeName -eq $timerServiceInstanceName)
            {
                [string]$serverName = $server.Name
 
                Write-Host -foregroundcolor DarkGray -NoNewline "Stop '$timerServiceName' service on server: "
                Write-Host -foregroundcolor Gray $serverName
 
                $service = Get-WmiObject -ComputerName $serverName Win32_Service -Filter "DisplayName='$timerServiceName'"
                $serviceInternalName = $service.Name
                sc.exe \\$serverName stop $serviceInternalName > $null
 
                # Wait until this service has actually stopped
                WaitForServiceState $serverName $timerServiceName "Stopped"
                 
                break;
            }
        }
    }
 
    Write-Host ""
}
 
#<summary>
# Waits for the service on the server to reach the required service state.
# This can be used to wait for the "SharePoint 2010 Timer" service to stop or to start
#</summary>
#<param name="$serverName">The name of the server with the service to monitor.</param>
#<param name="$serviceName">The name of the service to monitor.</param>
#<param name="$serviceState">The service state to wait for, e.g. Stopped, or Running.</param>
function WaitForServiceState([string]$serverName, [string]$serviceName, [string]$serviceState)
{
    Write-Host -foregroundcolor DarkGray -NoNewLine "Waiting for service '$serviceName' to change state to $serviceState on server $serverName"
 
    do
    {
        Start-Sleep 1
        Write-Host -foregroundcolor DarkGray -NoNewLine "."
        $service = Get-WmiObject -ComputerName $serverName Win32_Service -Filter "DisplayName='$serviceName'"
    }
    while ($service.State -ne $serviceState)
 
    Write-Host -foregroundcolor DarkGray -NoNewLine " Service is "
    Write-Host -foregroundcolor Gray $serviceState
}
 
#<summary>
# Starts the SharePoint Timer Service on each server in the SharePoint Farm.
#</summary>
#<param name="$farm">The SharePoint farm object.</param>
function StartSharePointTimerServicesInFarm([Microsoft.SharePoint.Administration.SPFarm]$farm)
{
    Write-Host ""
     
    # Iterate through each server in the farm, and each service in each server
    foreach($server in $farm.Servers)
    {
        foreach($instance in $server.ServiceInstances)
        {
            # If the server has the timer service then start the service
            if($instance.TypeName -eq $timerServiceInstanceName)
            {
                [string]$serverName = $server.Name
 
                Write-Host -foregroundcolor DarkGray -NoNewline "Start '$timerServiceName' service on server: "
                Write-Host -foregroundcolor Gray $serverName
 
                $service = Get-WmiObject -ComputerName $serverName Win32_Service -Filter "DisplayName='$timerServiceName'"
                [string]$serviceInternalName = $service.Name
                sc.exe \\$serverName start $serviceInternalName > $null
 
                WaitForServiceState $serverName $timerServiceName "Running"
                 
                break;
            }
        }
    }
 
    Write-Host ""
}
 
#<summary>
# Removes all xml files recursive on an UNC path
#</summary>
#<param name="$farm">The SharePoint farm object.</param>
function DeleteXmlFilesFromConfigCache([Microsoft.SharePoint.Administration.SPFarm]$farm)
{
    Write-Host ""
    Write-Host -foregroundcolor DarkGray "Delete xml files"
 
    [string] $path = ""
 
    # Iterate through each server in the farm, and each service in each server
    foreach($server in $farm.Servers)
    {
        foreach($instance in $server.ServiceInstances)
        {
            # If the server has the timer service delete the XML files from the config cache
            if($instance.TypeName -eq $timerServiceInstanceName)
            {
                [string]$serverName = $server.Name
                 
                Write-Host -foregroundcolor DarkGray -NoNewline "Deleting xml files from config cache on server: "
                Write-Host -foregroundcolor Gray $serverName
 
                # Remove all xml files recursive on an UNC path
                $path = "\\" + $serverName + "\c$\ProgramData\Microsoft\SharePoint\Config\*-*\*.xml"
                Remove-Item -path $path -Force
                 
                break
            }
        }
    }
 
    Write-Host ""
}
 
#<summary>
# Clears the SharePoint cache on an UNC path
#</summary>
#<param name="$farm">The SharePoint farm object.</param>
function ClearTimerCache([Microsoft.SharePoint.Administration.SPFarm]$farm)
{
    Write-Host ""
    Write-Host -foregroundcolor DarkGray "Clear the cache"
 
    [string] $path = ""
 
    # Iterate through each server in the farm, and each service in each server
    foreach($server in $farm.Servers)
    {
        foreach($instance in $server.ServiceInstances)
        {
            # If the server has the timer service then force the cache settings to be refreshed
            if($instance.TypeName -eq $timerServiceInstanceName)
            {
                [string]$serverName = $server.Name
 
                Write-Host -foregroundcolor DarkGray -NoNewline "Clearing timer cache on server: "
                Write-Host -foregroundcolor Gray $serverName
 
                # Clear the cache on an UNC path
                # 1 = refresh all cache settings
                $path = "\\" + $serverName + "\c$\ProgramData\Microsoft\SharePoint\Config\*-*\cache.ini"
                Set-Content -path $path -Value "1"
                 
                break
            }
        }
    }
 
    Write-Host ""
}
 
#**************************************************************************************
# Main script block
#**************************************************************************************
 
# Get the local farm instance
[Microsoft.SharePoint.Administration.SPFarm]$farm = [Microsoft.SharePoint.Administration.SPFarm]::get_Local()
 
# Stop the SharePoint Timer Service on each server in the farm
StopSharePointTimerServicesInFarm $farm
 
# Delete all xml files from cache config folder on each server in the farm
DeleteXmlFilesFromConfigCache $farm
 
# Clear the timer cache on each server in the farm
ClearTimerCache $farm
 
# Start the SharePoint Timer Service on each server in the farm
StartSharePointTimerServicesInFarm $farm