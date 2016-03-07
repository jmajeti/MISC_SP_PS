Write-Host -foregroundcolor Green "Restarting IIS on all the servers in FARM..."

[void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][reflection.assembly]::LoadWithPartialName("Microsoft.SharePoint.Administration")
[void][reflection.assembly]::LoadWithPartialName("System")
[void][reflection.assembly]::LoadWithPartialName("System.Collections")

[Microsoft.SharePoint.Administration.SPFarm]$farm = [Microsoft.SharePoint.Administration.SPFarm]::get_Local()

foreach ($server in $farm.Servers)
{
    if($server.Role -ne "Invalid")
    {
       Write-Host -foregroundcolor White ""
       Write-Host -foregroundcolor Yellow "Restarting IIS on server $server..."
       IISRESET $server.Name /noforce
       Write-Host -foregroundcolor Yellow "IIS status for server $server"
       IISRESET $server.Name /status
    }
}
Write-Host Write-Host -foregroundcolor Green IIS has been restarted on all servers
Read-Host 'Done...'
