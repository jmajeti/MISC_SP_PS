Add-PSSnapin Microsoft.SharePoint.Powershell
#To Email Address
$email="Username@gmail.com";
#Email BODY
$body="sample email body";
$subject="Test Subject";
#Get Web Instance
$web=Get-SPWeb "http://servername:8500";
#Calling SendMail() Method with Arguments
[Microsoft.SharePoint.Utilities.SPUtility]::SendEmail($web,0,0,$email,$subject,$body);