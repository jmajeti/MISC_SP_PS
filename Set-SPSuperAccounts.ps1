#################################################################################
# Powershell script to set the SuperUser and SuperReader accounts in SharePoint #
#                                      by                                       #
#                             Cornelius J. van Dyk                              #
# Blog: http://blog.cjvandyk.com    G+: http://aurl.to/+     Twitter: @cjvandyk #
#################################################################################

param([string]$cacheSuperAccount= "NIMHSP13SuperUserD",[string]$cacheReaderAccount= "NIH\NIMHSP13SuperReaderD")   
write-host ""
write-host -f White "Configure the WebApp property: portalsuperuseraccount and portalsuperreaderaccount"  
write-host ""
$snapin="Microsoft.SharePoint.PowerShell"
if (get-pssnapin $snapin -ea "silentlycontinue") 
{
  write-host -f Green "PSsnapin $snapin is loaded"
} 
else 
{     
  if (get-pssnapin $snapin -registered -ea "silentlycontinue") 
  {         
    write-host -f Green "PSsnapin $snapin is registered"        
    Add-PSSnapin $snapin        
    write-host -f Green "PSsnapin $snapin is loaded"    
  }     
  else 
  {         
    write-host -f Red "PSSnapin $snapin not found"    
  } 
}   
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")   
write-host -f Green "Getting current Farm"
$farm = Get-SPFarm  
$cacheSuperAccountCL = New-SPClaimsPrincipal -Identity $cacheSuperAccount -IdentityType WindowsSamAccountName 
$cacheReaderAccountCL= New-SPClaimsPrincipal -Identity $cacheReaderAccount -IdentityType WindowsSamAccountName 
$cacheSuperAccountCL = $cacheSuperAccountCL.ToEncodedString() 
$cacheReaderAccountCL= $cacheReaderAccountCL.ToEncodedString()   
write-host ""
write-host -f Green "Looping Web Applications"  
Get-SPWebApplication | foreach-object {        
  write-host ""      
  if ($_.UseClaimsAuthentication) 
  {             
    write-host -f white $_.Url " is a Claims Based Authentication WebApp"          
    write-host -f yellow " - Setting Policy: $cacheSuperAccountCL to Full Control for WebApp" $_.Url         
    $policy1 = $_.Policies.Add($cacheSuperAccountCL ,$cacheSuperAccount)         
    $policy1.PolicyRoleBindings.Add($_.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl))            
    write-host -f yellow " - Setting Property: portalsuperuseraccount $cacheSuperAccountCL for" $_.Url         
    $_.Properties["portalsuperuseraccount"] = $cacheSuperAccountCL          
    write-host -f yellow " - Setting Policy: $cacheReaderAccountCL to Full Read for WebApp" $_.Url         
    $policy2 = $_.Policies.Add($cacheReaderAccountCL ,$cacheReaderAccount)         
    $policy2.PolicyRoleBindings.Add($_.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead))            
    write-host -f yellow " - Setting Property: portalsuperreaderaccount $cacheReaderAccountCL for" $_.Url         
    $_.Properties["portalsuperreaderaccount"] = $cacheReaderAccountCL    
  }     
  else  
  {         
    write-host -f white $_.Url " is a Classic Authentication WebApp"          
    write-host -f yellow " - Setting Policy: $cacheSuperAccount to Full Control for WebApp" $_.Url         
    $policy1 = $_.Policies.Add($cacheSuperAccount ,$cacheSuperAccount)         
    $policy1.PolicyRoleBindings.Add($_.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl))            
    write-host -f yellow " - Setting Property: portalsuperuseraccount $cacheSuperAccount for" $_.Url         
    $_.Properties["portalsuperuseraccount"] = "$cacheSuperAccount"          
    write-host -f yellow " - Setting Policy: $cacheReaderAccount to Full Read for WebApp" $_.Url         
    $policy2 = $_.Policies.Add($cacheReaderAccount ,$cacheReaderAccount )         
    $policy2.PolicyRoleBindings.Add($_.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead))            
    write-host -f yellow " - Setting Property: portalsuperreaderaccount $cacheReaderAccount for" $_.Url         
    $_.Properties["portalsuperreaderaccount"] = "$cacheReaderAccount"    
  }       
  $_.Update()     
  write-host "Saved properties"
}   
Write ""
Write-host -f red "Runing IISReset"
IISreset /noforce 
Write "" 
