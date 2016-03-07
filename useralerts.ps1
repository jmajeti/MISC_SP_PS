foreach ($site in Get-SPSite -Limit All) 
{ 
  "Site Collection $site" 
  foreach ($web in $site.allwebs)
  {
     "  Web $web"
     $c = $web.alerts.count
     "    Deleting $c alerts"
     for ($i=$c-1;$i -ge 0; $i--) { $web.alerts.delete($i) }
  }
}