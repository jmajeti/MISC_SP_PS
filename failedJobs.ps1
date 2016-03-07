param([Int]$days,[String]$path) 
$f = get-spfarm 
$ts = $f.TimerService 
$jobs = $ts.JobHistoryEntries | ?{$_.Status -eq "Failed" -and $_.StartTime -gt ((get-date).AddDays(-$days))}  
 
$items = New-Object psobject 
$items | Add-Member -MemberType NoteProperty -Name "Title" -value "" 
$items | Add-Member -MemberType NoteProperty -Name "Server" -value "" 
$items | Add-Member -MemberType NoteProperty -Name "Status" -value "" 
$items | Add-Member -MemberType NoteProperty -Name "StartTime" -value "" 
$items | Add-Member -MemberType NoteProperty -Name "EndTime" -value "" 
$items | Add-Member -MemberType NoteProperty -Name "Duration" -value "" 
$a = $null 
$a = @() 
 
foreach($i in $jobs) 
{ 
$b = $items | Select-Object *;  
$b.Title = $i.JobDefinitionTitle; 
$b.Server = $i.ServerName; 
$b.Status = $i.Status; 
$b.StartTime = $i.StartTime; 
$b.EndTime = $i.EndTime; 
$b.Duration = ($i.EndTime - $i.StartTime); 
$a += $b; 
} 
$a | Where-Object {$_} | Export-Csv -Delimiter "," -Path c:\failedjobs.csv







function sendMail{

     #Write-Host "Sending Email"

     $file = "c:\failedjobs.csv"

     $att = new-object Net.Mail.Attachment($file)

     #SMTP server name
     $smtpServer = "210.150.176.100"

     #Creating a Mail object
     $msg = new-object Net.Mail.MailMessage

     #Creating SMTP server object
     $smtp = new-object Net.Mail.SmtpClient($smtpServer)

     #Email structure 
     $msg.From = "noreply-CorpQA@asddhg.com"

    
     $msg.CC.Add("prasad.tandel@asdfgh.com")

      $msg.To.Add("PJadhao@asdfg.com")

     $msg.subject = "Failed Jobs of Corporate Portal"

     $msg.body = "Hi Team, Please find Attachment"

     $msg.Attachments.Add($att)

     #Sending email 

     $smtp.Send($msg)
     $att.Dispose()
}

sendMail





