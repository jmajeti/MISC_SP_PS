[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
$url = $args[0];
$site = new-object microsoft.sharepoint.spsite($url);
for ($i=0;$i -lt $site.allwebs.count;$i++)
{
  write-host $site.allwebs[$i].url "...deleting" $site.allwebs[$i].recyclebin.count "item(s).";
  $site.allwebs[$i].recyclebin.deleteall();
}
write-host $site.url "...deleting" $site.recyclebin.count "item(s).";
$site.recyclebin.deleteall();
$site.dispose();
