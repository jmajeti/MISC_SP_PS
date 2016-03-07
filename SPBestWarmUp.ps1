#==================================================================================================================
#  http://spbestwarmup.codeplex.com/
#==================================================================================================================
#  Filename:        SPBestWarmUp.ps1
#  Author:          Jeff Jones
#  Version:         1.3
#  Last Modified:   10-15-2013
#  Description:     Warmup IIS memory cache by viewing pages from Internet Explorer.  Loads the full page so 
#                   resources like CSS, JS, and images are included.  Please modify lines 75-92 to suit your
#                   portal content design (popular URLs, custom pages, etc.)
#
#                   Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
#
#==================================================================================================================

param (
	[switch]$install
)

Function Installer() {
	# Add to Task Scheduler
	Write-Host "  Installing to Task Scheduler..." -ForegroundColor Green
	$user = $ENV:USERDOMAIN+"\"+$ENV:USERNAME
	Write-Host "  Current User: $user"
	
	# Attempt to detect password from IIS Pool (if current user is local admin & farm account)
	$appPools = gwmi -namespace "root\MicrosoftIISV2" -class "IIsApplicationPoolSetting" | select WAMUserName, WAMUserPass
	foreach ($pool in $appPools) {			
		if ($pool.WAMUserName -like $user) {
			$pass = $pool.WAMUserPass
			if ($pass) {
				break
			}
		}
	}
	
	# Manual input if auto detect failed
	if (!$pass) {
		$pass = Read-Host "Enter password for $user "
	}
	
	# Create Task
	schtasks /create /tn "SPBestWarmUp" /ru $user /rp $pass /rl highest /sc daily /st 01:00 /ri 60 /du 24:00 /tr "PowerShell.exe -ExecutionPolicy Bypass $global:path"
	Write-Host "  [OK]" -ForegroundColor Green
	Write-Host
}

Function WarmUp() {
	# Get URL list
	Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
	$was = Get-SPWebApplication -IncludeCentralAdministration
	$was |? {$_.IsAdministrationWebApplication -eq $true} |% {$caTitle = Get-SPWeb $_.Url | select Title}
	
	# Warmup SharePoint web applications
	Write-Host "Opening Web Applications..."
	$global:ie = New-Object -com "InternetExplorer.Application"
	$global:ie.Navigate("about:blank")
	$global:ie.visible = $true
	$global:ieproc = (Get-Process -Name iexplore)| Where-Object {$_.MainWindowHandle -eq $global:ie.HWND}
	foreach ($wa in $was) {
		$url = $wa.Url
		IENavigateTo $url
		IENavigateTo $url"_layouts/viewlsts.aspx"
		IENavigateTo $url"_vti_bin/UserProfileService.asmx"
		IENavigateTo $url"_vti_bin/sts/spsecuritytokenservice.svc"
	}
	
	# Warmup custom URLs
	Write-Host "Opening Custom URLs..."
	IENavigateTo "http://localhost:32843/Topology/topology.svc"
	
	# Add your own URLs below.  Looks at Central Admin Site Title for full lifecycle support in a single script file.
	switch -wildcard ($caTitle) {
		"*PROD*" {
			#IENavigateTo "http://portal/popularPage.aspx"
			#IENavigateTo "http://portal/popularPage2.aspx"
			#IENavigateTo "http://portal/popularPage3.aspx
		}
		"*TEST*" {
			#IENavigateTo "http://portal/popularPage.aspx"
			#IENavigateTo "http://portal/popularPage2.aspx"
			#IENavigateTo "http://portal/popularPage3.aspx
		}
		"*DEV*" {
			#IENavigateTo "http://portal/popularPage.aspx"
			#IENavigateTo "http://portal/popularPage2.aspx"
			#IENavigateTo "http://portal/popularPage3.aspx
		}
		default {
			#IENavigateTo "http://portal/popularPage.aspx"
			#IENavigateTo "http://portal/popularPage2.aspx"
			#IENavigateTo "http://portal/popularPage3.aspx
		}
	}
	
	# Warmup Host Name Site Collections (HNSC)
	Write-Host "Opening Host Name Site Collections (HNSC)..."
	$hnsc = Get-SPSite -Limit All |? {$_.HostHeaderIsSiteName -eq $true} | select Url
	foreach ($sc in $hnsc) {
		IENavigateTo $sc.Url
	}
	
	# Close IE window
	if ($global:ie) {
		Write-Host "Closing IE"
		$global:ie.Quit()
	}
	
	# Clean Temporary Files
	Remove-item "$env:systemroot\system32\config\systemprofile\appdata\local\microsoft\Windows\temporary internet files\content.ie5\*.*" -Recurse -ErrorAction SilentlyContinue
	Remove-item "$env:systemroot\syswow64\config\systemprofile\appdata\local\microsoft\Windows\temporary internet files\content.ie5\*.*" -Recurse -ErrorAction SilentlyContinue
}

Function IENavigateTo([string] $url, [int] $delayTime = 500) {
	# Navigate to a given URL
	Write-Host "  Navigating to $url"
	try {
		$global:ie.Navigate($url)
	} catch {
		$pid = $global:ieproc.id
		Write-Host "  IE not responding.  Closing process ID $pid"
		$global:ie.Quit()
		$global:ieproc | Stop-Process -Force
		$global:ie = New-Object -com "InternetExplorer.Application"
		$global:ie.Navigate("about:blank")
		$global:ie.visible = $true
		$global:ieproc = (Get-Process -Name iexplore)| Where-Object {$_.MainWindowHandle -eq $global:ie.HWND}
	}
	IEWaitForPage $delayTime
}

Function IEWaitForPage([int] $delayTime = 500) {
	# Wait for current page to finish loading
	$loaded = $false
	$loop = 0
	$maxLoop = 20
	while ($loaded -eq $false) {
		$loop++
		if ($loop -gt $maxLoop) {
			$loaded = $true
		}
		[System.Threading.Thread]::Sleep($delayTime) 
		# If the browser is not busy, the page is loaded
		if (-not $global:ie.Busy)
		{
			$loaded = $true
		}
	}
}

#Main
Write-Host "SPBestWarmUp v1.3  (last updated 10-15-2013)`n"

#Check Permission Level
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Break
} else {
    #Warmup
    $global:path = $MyInvocation.MyCommand.Path
    $tasks = schtasks /query /fo csv | ConvertFrom-Csv
    $spb = $tasks | Where-Object {$_.TaskName -eq "\SPBestWarmUp"}
    if (!$spb -and !$install) {
	    Write-Host "Tip: to install on Task Scheduler run the command ""SPBestWarmUp.ps1 -install""" -ForegroundColor Yellow
    }
    if ($install) {
	    Installer
    }
    WarmUp
}