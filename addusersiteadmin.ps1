# This script will add a named Site Collection Administrator
# to all Site Collections within a Web Application.
#
######################## Start Variables ########################
$newSiteCollectionAdminLoginName = "nih\hemchandraf"
$newSiteCollectionAdminEmail = "fnu.hemchandra@nih.gov"
$newSiteCollectionAdminName = "Hem Chandra, Fnu"
$newSiteCollectionAdminNotes = ""
$siteURL = "https://share.nibib.nih.gov/SitePages/Home.aspx" 
$add = 1
######################## End Variables ########################
Clear-Host
$siteCount = 0
[system.reflection.assembly]::loadwithpartialname("Microsoft.SharePoint")
$site = new-object microsoft.sharepoint.spsite($siteURL)
$webApp = $site.webapplication
$allSites = $webApp.sites
######################## Write Progress Declaration ######################## 
$i = 0
foreach ($site in $allSites)
{
    $web = $site.openweb()
    $web.allusers.add($newSiteCollectionAdminLoginName, $newSiteCollectionAdminEmail, $newSiteCollectionAdminName, $newSiteCollectionAdminNotes)

    $user = $web.allUsers[$newSiteCollectionAdminLoginName]
    $user.IsSiteAdmin = $add
    $user.Update()
    $web.Dispose()
    $siteCount++

######################## Update Counter and Write Progress ########################
   $i++
   Write-Progress -Activity "Adding $newSiteCollectionAdminName to all site collections within $siteURL.  Please wait..." -status "Added: $i of $($allSites.Count)" -percentComplete (($i / $allSites.Count)  * 100)
}
$site.dispose()
write-host "Updated" $siteCount "Site Collections."