# Get content
$PatchCode = "KB958644" #Patch version to test
$computers = get-content C:\ServersList.txt #TextFile with all

[String]$mylogin = "DOMAIN\AdminLogin" #Admin login
[String]$mypassword = "xxxxx" #Admin Password
$mySecurePW = new-object Security.SecureString
$mypassword.ToCharArray()|% { $mySecurePW.AppendChar($_) }

# Create the Credential used
$myCredential = New-Object System.Management.Automation.PsCredential($mylogin, $mySecurePW)



# Get all the info using WMI
foreach($myServer in $computers)
{
	write-host "--------------------------------------"
	write-host $myServer
	$results = get-wmiobject -class "Win32_QuickFixEngineering" -namespace "root\CIMV2" -computername $myServer -Credential $myCredential

	# Loop through $results and look for a match then output to screen 
	foreach ($objItem in $results) 
	{ 
	    if ($objItem.HotFixID -match $PatchCode) 
	    {
	        write-host "Hotfix $PatchCode installed"
	    }
	}
	write-host "--------------------------------------"
}
