#==================================================================================================================
#  http://comb.codeplex.com/
#==================================================================================================================
#  Filename:        EventComb.ps1
#  Author:          Jeff Jones
#  Version:         1.0
#  Last Modified:   09-26-2013
#  Description:     Gather eventlogs from servers into a daily summary email with attached CSV detail.
#                   Helps administrators be proactive with issue resolution by better understanding internal 
#                   server health.  Open with Microsoft Excel for PivotTable, PivotCharts, and further analysis.
#
#					NOTE - Please adjust lines 26-33 for your environment.
#
#                   Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
#
#==================================================================================================================

param (
	[switch]$install
)

# Configuration
$global:configHours = -24											# time threshold (previous day)
$global:configMaxEvents = 999										# maximum number of events from any 1 server
$global:configTargetMachines = @("spautodetect")					# servers  to gather eventlogs from.  use = @("spautodetect") for SharePoint farm auto detection
$global:configSendMailTo = @("admin1@demo.com","admin2@demo.com")	# to address
$global:configSendMailFrom = "no-reply@demo.com"					# from address
$global:configSendMailHost = "mailserver"							# outbound SMTP mail server
$global:configWarnDisk = 0.25										# threshold for warning (25%)
$global:configErrorDisk = 0.10										# threshold for warning (10%)
$global:configExcludeMaintenanceHours = @(21,22,23,0,1,2,3)			# exclude 11PM-4AM nightly maintenance window
$global:configExcludeEventSources = @("McAfee PortalShield~2053")	# exclude known event sources

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
	schtasks /create /tn "EventComb" /ru $user /rp $pass /rl highest /sc daily /st 03:00 /tr "PowerShell.exe -ExecutionPolicy Bypass $global:path"
	Write-Host "  [OK]" -ForegroundColor Green
	Write-Host
}

Function EventComb() {
	# Auto detect on SharePoint farms
	if ($global:configTargetMachines -eq @("spautodetect")) {
		$global:configTargetMachines = @()
		foreach ($s in ((Get-SPFarm).Servers |? {$_.Role -ne "Invalid"} )) {
			$global:configTargetMachines += $s.Address
		}
	}
	
	# Initialize
	$start = Get-Date
	$logAfter = (Get-Date).AddHours($global:configHours)
	Write-Host ("{0} machine(s) targeted" -f $global:configTargetMachines.Count)
	$csv = @()

	# Loop for all machines
	foreach ($machine in $global:configTargetMachines) {
		foreach ($log in @("Application", "System")) {
			 Write-Host ("Gathering log {0} for {1} ... " -f $log, $machine) -NoNewline
			 # Gather event log detail
			 foreach ($type in @("Error","Warning")) {
				 $events = Get-EventLog -ComputerName $machine -Logname $log -After $logAfter -EntryType $type -Newest $global:configMaxEvents
				 if ($events) {
					foreach ($e in $events) {
						$keep = $true
						# Exclude based on ID and Source
						foreach ($skip in $global:configExcludeEventSources) {
							if ($e.Source -eq  $skip.Split("~")[0] -and $e.InstanceID -eq $skip.Split("~")[1]) {
								$keep = $false
							}
						}
						# Exclude based on maintenance hours
						foreach ($hour in $global:configExcludeMaintenanceHours) {
							if ($e.TimeWritten.Hour -eq $hour) {
								$keep = $false
							}
						}
						# Append to CSV
						if ($keep) {
							$csv += $e
						}
					 }
				 }
			 }
			 Write-Host "[OK]" -ForegroundColor Green
		}
	}

	# Write CSV file
	Write-Host "Writing CSV file ..." -NoNewline
	$csv | Export-Csv -Path "EventComb.csv" -NoTypeInformation -Force
	Write-Host "[OK]" -ForegroundColor Green

	# Format HTML summary
	$totalErr = 0
	$totalWarn = 0
	$html = ("The below table summaries the eventlog entries of the last {0} hours on these machines:<br><br><table>" -f $global:configHours)
	$html += "<tr><td>&nbsp;</td><td width='20px'>&nbsp;</td><td><b>Error</b></td><td></td><td><b>Warn</b></td></tr>"
	foreach ($machine in $global:configTargetMachines) {
		# Summary total for Errors
		$countErr = 0
		$logErr = ($csv |? {$_.EntryType -eq "Error" -and $_.MachineName -like "$machine*"})
		if ($logErr) {
			$countErr = $logErr.Count
			$totalErr += $countErr
		}
		# Summary total for Warnings
		$countWarn = 0
		$logWarn = ($csv |? {$_.EntryType -eq "Warning" -and $_.MachineName -like "$machine*"})
		if ($logWarn) {
			$countWarn = $logWarn.Count
			$totalWarn += $countWarn
		}
		$html += ("<tr><td>{0}</td><td>&nbsp;</td><td style='background-color: #FF9D9D;'>{1}</td><td style='color: #FFFFFF'>-</td><td style='background-color: #FFFF6C;'>{2}</td></tr>" -f $machine, $countErr, $countWarn)
	}
	$html += ("<tr><td>&nbsp;</td><td width='20px'>&nbsp;</td><td>{0}</td><td style='color: #FFFFFF'>-</td><td>{1}</td></tr>" -f $totalErr, $totalWarn)
	$html += "</table>"

	# Format HTML pivot tables
	Write-Host "Pivot tables ... " -NoNewline
	$html += "<br><br><table><tr><td colspan=3><b>Source Pivots</b></td></tr>"
	$groups = $csv | group Source -NoElement | sort Count -Descending
	foreach ($g in $groups) {
		$html += "<tr><td> " + $g.Name + " </td><td style='color: #FFFFFF'>-</td><td style='background-color: #99FF99;'> " + $g.Count + " </td></tr>";
	}
	$html += "</table>"
	$html += "<br><br><table><tr><td colspan=3><b>EventID Pivots</b></td></tr>"
	$groups = $csv | group EventID -NoElement | sort Count -Descending
	foreach ($g in $groups) {
		$html += "<tr><td> " + $g.Name + " </td><td style='color: #FFFFFF'>-</td><td style='background-color: #99CCFF;'> " + $g.Count + " </td></tr>";
	}
	$html += "</table>"
	Write-Host "[OK]" -ForegroundColor Green

	# Free disk space
	$html += "<br><br><table><tr><td colspan=3><b>Free Disk</b></td></tr>"
	foreach ($machine in $global:configTargetMachines) {
		$html += "<tr><td valign='top'>$machine</td></tr>"
		$wql = "SELECT Size, FreeSpace, Name, FileSystem FROM Win32_LogicalDisk WHERE DriveType = 3"
		$wmi = Get-WmiObject -ComputerName $machine -Query $wql
		foreach ($w in $wmi) {
			$color = "#FFFFFF"
			$note = ""
			$letter = $w.Name
			$freeSpace = ($w.FreeSpace / 1GB)
			$prctFree = ($w.FreeSpace / $w.Size)
			if ($prctFree -lt $global:configWarnDisk) {
				$color = "#FFFF6C"
			}
			if ($prctFree -lt $global:configErrorDisk) {
				$color = "#FF9D9D"
				$note = "*"
			}
			$html += ("<tr><td></td><td>$letter</td><td style='background-color: $color;'>&nbsp;{0:N1} GB ({1:P0}) $note&nbsp;</td></tr>" -f $freeSpace, $prctFree)
		}
	}
	$html += "</table>"

	# Send email summary with CSV attachment
	Send-MailMessage -To $global:configSendMailTo -From $global:configSendMailFrom -Subject "EventComb" -BodyAsHtml -Body $html -Attachments "EventComb.csv" -SmtpServer $global:configSendMailHost
	Write-Host ("Operation completed successfully in {0} seconds" -f ((Get-Date) - $start).Seconds)
}

#Main
Write-Host "EventComb v1.0  (last updated 09-26-2013)`n"


#Check Permission Level
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Break
} else {
    #EventComb
    $global:path = $MyInvocation.MyCommand.Path
    $tasks = schtasks /query /fo csv | ConvertFrom-Csv
    $spb = $tasks | Where-Object {$_.TaskName -eq "\EventComb"}
    if (!$spb -and !$install) {
	    Write-Host "Tip: to install on Task Scheduler run the command ""EventComb.ps1 -install""" -ForegroundColor Yellow
    }
    if ($install) {
	    Installer
    }
    EventComb
}