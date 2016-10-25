Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
 
#Starting web app
$site = "https://sharepoint.hrsa.gov"
 
# Function: FindAccessEmail
# Description: Go through a target web application and list the title, url and access request email.
function FindAccessEmail
{
	$WebApps = Get-SPWebApplication($site)
	foreach($WebApplication in $WebApps)
	{
	    foreach ($Collection in $WebApplication.Sites)
	    {
	       foreach($Web in $Collection.AllWebs)
	        {
				$siteDetails = $Web.title+'#'+$Web.url+'#'+$Web.RequestAccessEmail 
	            write-host $siteDetails
				Write-Output $siteDetails
	        }
	    }
	  }
}
#Run Script!
FindAccessEmail