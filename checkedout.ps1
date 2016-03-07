[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
 
function CheckedOutItems() {
 
write-host "Please enter the site url"
$url = read-host
write ("SiteURL`t" + "FileName`t" +  "CheckedOutTo`t" + "ModifiedDate`t"+"Version")
$site = New-Object Microsoft.SharePoint.SPSite($url)
$webs = $site.AllWebs
foreach($web in $webs)
{
$listCollections = $web.Lists
foreach($list in $listCollections)
{
 
 
if ($list.BaseType.ToString() -eq "DocumentLibrary")
{
 $dList = [Microsoft.Sharepoint.SPDocumentLibrary]$list
 $items = $dList.Items
$files = $dList.CheckedOutFiles
foreach($file in $files)
{
 
$wuse = $file.DirName.Substring($web.ServerRelativeUrl.Length)
Write ($web.Url+ "`t" + $wuse+"`/" + $file.LeafName +  "`t" + $file.CheckedOutBy.Name + "`t" + $file.TimeLastModified.ToString()+"`t" + "No Checked In Version" )
}
 foreach($item in $items)
 {
 if ($item["Checked Out To"] -ne $null)
  {
$splitStrings = $item["Checked Out To"].ToString().Split('#')
 
  Write ($web.Url+ "`t" + $item.Url + "`t" + $splitStrings[1].ToString() + "`t" + $item["Modified"].ToString() +"`t" + $item["Version"].ToString())
 }
 }
 
  
}
 
 
 
}
$web.Dispose()
}
$site.Dispose()
}
 
 
CheckedOutItems
