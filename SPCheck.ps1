# SHAREPOINT SERVER REPORT V1.1
# 3/7/2015
# This script gets basic farm server information, including
# information on: Key system services (eg, AppFabric, FIM), 
# SharePoint Services, IIS sites and app pools, App Fabric Status
# and so on, writing the results to a text files, and then
# opening that text file.
# 03/11/15: added AppFabric Health section
#################################################################

# Install necessary modules, snapins and features

Import-Module webadministration
If ((Get-PSSnapin -Name "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null ){Add-PsSnapin Microsoft.SharePoint.PowerShell}
If ((Get-WindowsFeature Web-WMI).installed -eq "False" ) {Add-WindowsFeature Web-WMI}

# Set global variables

$ComputerName = $env:computername
$Date = Get-Date

# Write Report Header

$String = "SHAREPOINT SERVER SYSTEM CHECK | " + $ComputerName + " | " + $Date
Write-Host ""
Write-Host $String

Write-Host ""
Write-Host "Core SharePoint services"
$SPServices = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName | ? {($_.Name -eq "AppFabricCachingService") -or ($_.Name -eq "c2wts") -or ($_.Name -eq "FIMService") -or ($_.Name -eq "FIMSynchronizationService") -or ($_.Name -eq "Service Bus Gateway") -or ($_.Name -eq "Service Bus Message Broker") -or ($_.Name -eq "SPAdminV4") -or ($_.Name -eq "SPSearchHostController") -or ($_.Name -eq "OSearch15") -or ($_.Name -eq "SPTimerV4") -or ($_.Name -eq "SPTraceV4") -or ($_.Name -eq "SPUserCodeV4") -or ($_.Name -eq "SPWriterV4") -or ($_.Name -eq "FabricHostSvc") -or ($_.Name -eq "WorkflowServiceBackend")}
$SPServices | Select-Object DisplayName, StartName, StartMode, State, Status | Sort-Object DisplayName | Format-Table -AutoSize

Write-Host "IIS Sites"
$IIS = Get-ChildItem IIS:\Sites | Select-Object Name, ID, ApplicationPool, ServerAutoStart, State
$IIS | Format-Table -autosize

Write-Host "IIS application pools"
$AppPools = Get-CimInstance -Namespace root/MicrosoftIISv2 -ClassName IIsApplicationPoolSetting -Property Name, AppPoolIdentityType, WAMUserName, AppPoolState | 
select @{e={$_.Name.Replace('W3SVC/APPPOOLS/', '')};l="Name"}, AppPoolIdentityType, WAMUserName, AppPoolState
$AppPools | Format-Table -AutoSize

Write-Host "AppFabric Health"
$SPServices = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName | ? {($_.Name -eq "AppFabricCachingService")}

If (($SPServices.StartMode -eq "Auto") -AND ($SPServices.State -eq "Running")) 
{
Get-SPServiceInstance | ? {($_.service.tostring()) -eq "SPDistributedCacheService Name=AppFabricCachingService"} | select Server, Status | Format-Table -AutoSize
Use-CacheCluster
Get-CacheHost | Format-Table -AutoSize
}
Else
{
Write-Host ""
Write-Host "AppFabric disabled or not running on this machine"
Write-Host ""
}
Write-Host "Report completed"
Write-Host ""

