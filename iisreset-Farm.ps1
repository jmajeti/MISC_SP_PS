cls
Write-Host "Loading SharePoint Commandlets"
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
Write-Host -ForegroundColor Green " Commandlets Loaded ... Loading Variables"
Write-Host
[array]$servers= Get-SPServer | ? {$_.Role -eq "Application"}
$farm = Get-SPFarm
foreach ($server in $servers)
{
     Write-Host -ForegroundColor Yellow "Attempting to reset IIS for $server"
        iisreset $server /noforce "\\"$_.Address
        iisreset $server /status "\\"$_.Address
        Write-Host
        Write-Host -ForegroundColor Green "IIS has been reset for $server"
        Write-Host
}
Write-Host -ForegroundColor Green "IIS has been reset across the SharePoint Farm"
Start-Sleep -Seconds 5
Write-host