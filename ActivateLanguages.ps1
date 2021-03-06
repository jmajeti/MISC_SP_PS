if($args) {
    $siteUrl = $args[0] 
    
    $snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
    if ($snapin -eq $null)
    {
	   Write-Host "Loading Microsoft SharePoint Powershell Snapin"
	   Add-PSSnapin "Microsoft.SharePoint.Powershell"
    }

    $cultureinfo = new-object system.globalization.cultureinfo("fr-FR")
    
    $site = get-spsite $siteurl
    $site.allwebs | foreach { $_.IsMultiLingual = 'True' 
                              $_.AddSupportedUICulture($cultureinfo) 
                              $_.Update()
                            }
    
    $site = get-spsite $siteurl
    $site.allwebs |foreach {Write-Host $_.Title + " " +  $_.IsMultiLingual}
    
}
else
{
    Write-Host "ERROR: You must supply SiteCollection URL as parameter when calling ActivateLanguages.ps1"
}