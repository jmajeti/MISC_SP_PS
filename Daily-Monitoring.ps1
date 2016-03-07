# Author: Prasad Tandel
# Below script will help to reduce the time taken in monitoring daily Checklist


#Step 1: Get list of server from farm.
$Servers = Get-content C:\TEMP\servers.txt

# To get status of IIS app pool in Sharepoint farm
function ConvertTo-AppPoolState($value)
{
    switch($value)
    {
        1    {"Starting"}
        2    {"Started"}
        3    {"Stopping"}
        4    {"Stopped"}
        default    {"Unknown"}
    }
}
foreach ($server in $servers)
{
$AppPools =  [ADSI]"IIS://$server/W3SVC/AppPools"
$AppPools.Children | Select-Object Name,@{n='AppPoolState';e={ConvertTo-AppPoolState $_.AppPoolState.Value}},@{n='machinename';e={$server}} >> d:\Result.txt 
Write-Host -Separator `n
}

# To get status of SharePoint services and SQL services in farm
foreach ($server in $servers)
{
Get-Service SPAdminV4,OSearch14,SPTraceV4,SPTimerV4 -Computer $Server | Select name,DisplayName,status,machinename|format-table -autosize >> d:\Result.txt
}

# To get Diskspace status of all servers in farm
$DiskResults = @()
foreach ($computerName in $Servers)
{
# Write-Host $computerName
  $objDisks = Get-WmiObject -Computername $computerName -Class win32_logicaldisk | Where-Object { $_.DriveType -eq 3 }
  ForEach( $disk in $objDisks )
  {
    $diskFragmentation = "Unknown"
    try
    {
      $objDisk = Get-WmiObject -Computername $computerName -Class Win32_Volume -Filter "DriveLetter='$($disk.DeviceID)'"

    }
    catch{}
    $ThisVolume = "" | select ServerName,Volume,Capacity,FreeSpace
    $ThisVolume.ServerName = $computerName
    $ThisVolume.Volume = $disk.DeviceID
    $ThisVolume.Capacity = $([Math]::Round($disk.Size/1073741824,2))
    $ThisVolume.FreeSpace = $([Math]::Round($disk.FreeSpace/1073741824,2))
    $DiskResults += $ThisVolume
# Write-Host -Separator `n
  }
}
$DiskResults | ft -autosize >> d:\Result.txt

# To send mail on onstream support email id.

function sendMail{

     Write-Host "Sending Email"
     $file = "d:\Result.txt"
     $att = new-object Net.Mail.Attachment($file)
     #SMTP server name
     $smtpServer = "SMTP Server"
     #Creating a Mail object
     $msg = new-object Net.Mail.MailMessage
     #Creating SMTP server object
     $smtp = new-object Net.Mail.SmtpClient($smtpServer)
     #Email structure 
     $msg.From = "Emailid"
     $msg.ReplyTo = "Emailid; Emailid"
     $msg.To.Add("Emailid")
     $msg.subject = "Daily Checklist"
     $msg.body = ""
     $msg.Attachments.Add($att)
     #Sending email 
     $smtp.Send($msg)
     $att.Dispose()
}
#Calling function
sendMail

# End of Script
# Thanks 
#Param



