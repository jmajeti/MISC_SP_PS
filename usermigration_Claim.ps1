$WebAppName = "http://sps2013"

#THIS SHOULD BE THE ACCOUNT THAT IS RUNNING THIS SCRIPT, WHO SHOULD ALSO BE A LOCAL ADMIN
$account = "domain\administrator"

$wa = get-SPWebApplication $WebAppName


Set-SPwebApplication $wa -AuthenticationProvider (New-SPAuthenticationProvider) -Zone Default
# this will prompt about migration, CLICK YES and continue

#This step will set the admin for the site 
$wa = get-SPWebApplication $WebAppName
$account = (New-SPClaimsPrincipal -identity $account -identitytype 1).ToEncodedString()

#Once the user is added as admin, we need to set the policy so it can have the right access
$zp = $wa.ZonePolicies("Default")
$p = $zp.Add($account,"PSPolicy")
$fc=$wa.PolicyRoles.GetSpecialRole("FullControl")
$p.PolicyRoleBindings.Add($fc)
$wa.Update()

#Final step is to trigger the user migration process
$wa = get-SPWebApplication $WebAppName
$wa.MigrateUsers($true)
$wa.ProvisionGlobally()