#Function to Check if an User exists in AD
function CheckUserExistsInAD()
{
    Param( [Parameter(Mandatory=$true)] [string]$UserLoginID )
    
    #Search the User in AD
    $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    foreach ($Domain in $forest.Domains)
    {
        $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $Domain.Name)
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
    
        $root = $domain.GetDirectoryEntry()
        $search = [System.DirectoryServices.DirectorySearcher]$root
        $search.Filter = "(&(objectCategory=User)(samAccountName=$UserLoginID))"
        $result = $search.FindOne()
 
         if ($result -ne $null)
         {
           return $true
         }
    }
    return $false  
}
 
$WebAppURL="http://dev-apps:8080/"
$WebApp = Get-SPWebApplication $WebAppURL #Get all Site Collections of the web application
 

#Iterate through all Site Collections
foreach($site in $WebApp.Sites) 
{
    $Output = "C:\OrphanUserList.txt"
    $strOut = "User Name|User Login|URL"+"`r`n"

    #Get all Webs with Unique Permissions - Which includes Root Webs
    $WebsColl = $site.AllWebs | Where {$_.HasUniqueRoleAssignments -eq $True} | ForEach-Object {
         
    $OrphanedUsers = @()
         
    #Iterate through the users collection
    foreach($User in $_.SiteUsers)
    {
        #Exclude Built-in User Accounts , Security Groups & an external domain "corporate"
        if(($User.LoginName.ToLower() -ne "nt authority\authenticated users") -and
           ($User.LoginName.ToLower() -ne "sharepoint\system") -and
           ($User.LoginName.ToLower() -ne "nt authority\local service")  -and
           ($user.IsDomainGroup -eq $false ) #-and
                          #($User.LoginName.ToLower().StartsWith("corporate") -ne $true) 
          )
          {
            $UserName = $User.LoginName.split("\")  #Domain\UserName
            $AccountName = $UserName[1]    #UserName
            if ( ( CheckUserExistsInAD $AccountName) -eq $false )
            {
                Write-Host "$($User.Name)($($User.LoginName)) from $($_.URL) doesn't Exists in AD!"
                
		$strOut += $User.Name+"|"+$User.LoginName+"|"+$_.URL+"`r`n"
		$strOut|Out-File $Output
                     
                #Make a note of the Orphaned user
                $OrphanedUsers+=$User.LoginName
            }
          }
      }

	# ***********************************************************************************#
	#		  		               REMOVE ORPHAN USERS
	# ***********************************************************************************#

# Remove the Orphaned Users from the site
#    foreach($OrpUser in $OrphanedUsers)
#    {
#	   $_.SiteUsers.Remove($OrpUser)
#     Write-host "Removed the Orphaned user $($OrpUser) from $($_.URL) "
#    }
  }
}