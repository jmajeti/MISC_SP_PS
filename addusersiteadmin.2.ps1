# This script will add a named Site Collection Administrator
# to all Site Collections within a Web Application.
#
######################## Start Variables ########################
$newSiteCollectionAdminLoginName = "nih\nimhsp10cvt"
$newSiteCollectionAdminEmail = ""
$newSiteCollectionAdminName = "nimhsp10cvt"
$newSiteCollectionAdminNotes = ""
$siteURL = "https://nimhsharetst.nimh.nih.gov/" 
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