function Check-Permission_SPWeb([string]$webURL, [string]$myLoginToCheck)
{
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null
	$mysite = new-object Microsoft.SharePoint.SPSite($webURL)
	$myweb = $mysite.openweb()
	$mywebUsers = $myweb.Users

	foreach($mywebUser in $mywebUsers)
	{
		if($mywebUser.LoginName -eq $myLoginToCheck)
		{
			$myPermissions = $myweb.Permissions
			foreach($myPermission in $myPermissions)
			{
				if($mywebUser.ID -eq $myPermission.Member.ID)
				{
					Write-Host "      ----------------------------- "
					Write-Host " ---- WEB Url: ", $webURL
					Write-Host " ---- WEB Has Unique Permission:", $myweb.HasUniqueRoleAssignments
					Write-Host " ---- User: ", $mywebUser.LoginName, " - Permissions: ", $myPermission.PermissionMask.ToString()
					foreach ($role in $mywebUser.Roles)
					{
						if ($role.Type -ne [Microsoft.SharePoint.SPRoleType]::None)
						{
							Write-Host " ---- Role: ", $role.Type.ToString()
						}
					}
					Write-Host "      ----------------------------- "
				}
			}
			foreach($myList in $myweb.lists)
			{
				if($myList.HasUniqueRoleAssignments -eq $True)
				{
					$myListPermissions = $myList.Permissions
					foreach($myListPermission in $myListPermissions)
					{
						if($mywebUser.ID -eq $myListPermission.Member.ID)
						{
							Write-Host "            ----------------------------- "
							Write-Host " ---------- LIST NAME: ", $myList.Title
							Write-Host " ---------- LIST Has Unique Permission:", $myList.HasUniqueRoleAssignments
							Write-Host " ---------- User: ", $mywebUser.LoginName, " - Permissions: ", $myListPermission.PermissionMask.ToString()
							foreach ($roleAssignment in $myList.RoleAssignments)
							{
								if($mywebUser.ID -eq $roleAssignment.Member.ID)
								{
									foreach($mySPRoleDefinition in $roleAssignment.RoleDefinitionBindings)
									{
										if ($mySPRoleDefinition.Type -ne [Microsoft.SharePoint.SPRoleType]::None)
										{
											Write-Host " ---------- Role: ", $mySPRoleDefinition.Type.ToString()
										}
									}
								}
							}
							Write-Host "            ----------------------------- "
						}
					}
				}
			}
		}
	}
	foreach ($subweb in $myweb.GetSubwebsForCurrentUser())
	{
		Check-Permission_SPWeb $subweb.Url $myLoginToCheck
    }
	$myweb.Dispose()
	$mysite.Dispose()
}



function ListUsers([string]$SiteCollectionURL, [string]$LoginToCheck)
{
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null
	$site = new-object Microsoft.SharePoint.SPSite($SiteCollectionURL)
	$web = $site.openweb()

	$siteCollUsers = $web.SiteUsers

	foreach($MyUser in $siteCollUsers)
	{
		if($MyUser.LoginName -eq $LoginToCheck)
		{
			Write-Host " ------------------------------------- "
			Write-Host "Site Collection URL:", $SiteCollectionURL
			if($MyUser.IsSiteAdmin -eq $true)
			{
				Write-Host "ADMIN: ", $MyUser.LoginName
			}
			else
			{
				Write-Host "USER: ", $MyUser.LoginName
			}
			Check-Permission_SPWeb $web.Url $MyUser.LoginName
			Write-Host " ------------------------------------- "
		}
	}
	
	$web.Dispose()
	$site.Dispose()
}

function ListUsersForAllCollections([string]$WebAppURL, [string]$LoginToCheck)
{

	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") > $null

	$Thesite = new-object Microsoft.SharePoint.SPSite($WebAppURL)
	$oApp = $Thesite.WebApplication

	foreach ($Sites in $oApp.Sites)
	{
		$mySubweb = $Sites.RootWeb
		[string]$TempRelativeURL = $mySubweb.Url
		ListUsers $TempRelativeURL $LoginToCheck
    }
}

function StartProcess()
{
	# Create the stopwatch
	[System.Diagnostics.Stopwatch] $sw;
	$sw = New-Object System.Diagnostics.StopWatch
	$sw.Start()
	cls

	[string]$userlogintocheck = "DOMAIN\User_login"
	
	ListUsersForAllCollections "http://myWebApplication" $userlogintocheck

	$sw.Stop()

	# Write the compact output to the screen
	write-host "Login checked in Time: ", $sw.Elapsed.ToString()
}

StartProcess

