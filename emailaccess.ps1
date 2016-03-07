#Set-ExecutionPolicy RemoteSigned
$OutputFN = "AccessRequestConfigs.csv"
#delete the file, If already exist!
if (Test-Path $OutputFN)
 { 
    Remove-Item $OutputFN
 }
#Write the CSV Headers
"Site Collection Name, Site Name ,URL ,Access Requst E-Mail" > $OutputFN
 
# Get All Web Applications 
$WebAppServices=Get-SPWebApplication
 
foreach($webApp in $WebAppServices)
{
 
   # Get All Site collections
    foreach ($SPsite in $webApp.Sites)
    {
       # get All Sites 
       foreach($SPweb in $SPsite.AllWebs)
        {
          if($SPweb.RequestAccessEnabled -eq $True)
          {
           $SPsite.Rootweb.title + "," + $SPweb.title.replace(","," ") + "," + $SPweb.URL + "," + $SPweb.RequestAccessEmail >>$OutputFN
          
          }
          $SPweb.dispose()
        }
      $SPsite.dispose()
    }
  }


