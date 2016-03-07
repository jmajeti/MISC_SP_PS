function AddLoginAsSiteCollAdmin([string]$SiteCollectionURL, [string]$LoginNewAdmin)
{
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null
	$site = new-object Microsoft.SharePoint.SPSite($SiteCollectionURL)
	$web = $site.openweb()

	#Debugging - show SiteCollectionURL
	Write-Host "SiteCollectionURL", $SiteCollectionURL

	$siteCollUsers = $web.SiteUsers
	$siteCollUsers.Add($LoginNewAdmin, "", "", "")
	Write-Host "  ADMIN ADDED: ", $LoginNewAdmin
	$web.Update()
	$myuser = $siteCollUsers[$LoginNewAdmin]
	$myuser.IsSiteAdmin = $TRUE
	$myuser.Update()

	$web.Update()
	$web.Dispose()
	$site.Dispose()
}


function AddSiteCollAdminForAllCollections([string]$WebAppURL, [string]$LoginNewSiteAdmin)
{
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null
	$Thesite = new-object Microsoft.SharePoint.SPSite($WebAppURL)
	$oApp = $Thesite.WebApplication

	foreach ($Sites in $oApp.Sites)
	{
		$mySubweb = $Sites.RootWeb
		$TempRelativeURL = $mySubweb.Url
		AddLoginAsSiteCollAdmin $TempRelativeURL $LoginNewSiteAdmin
    }
}
function StartProcess()
{
	# Create the stopwatch
	[System.Diagnostics.Stopwatch] $sw;
	$sw = New-Object System.Diagnostics.StopWatch
	$sw.Start()
	cls
	$usertoreplaceinsiteadmin = "DOMAIN\SharePointAdmin_Login"
	
	AddSiteCollAdminForAllCollections "http://MyWebApplication1" $usertoreplaceinsiteadmin
	AddSiteCollAdminForAllCollections "http://MyWebApplication2" $usertoreplaceinsiteadmin

	$sw.Stop()

	# Write the compact output to the screen
	write-host $usertoreplaceinsiteadmin, " Login add as Site Collection Admin in Time: ", $sw.Elapsed.ToString()
}
StartProcess
