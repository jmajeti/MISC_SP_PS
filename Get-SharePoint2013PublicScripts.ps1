<#  
.SYNOPSIS  
       The purpose of this script is to download a set of SharePoint 2013 PowerShell scripts from the technet script repository  
.DESCRIPTION  
       The purpose of this script is to download a set of SharePoint 2013 PowerShell scripts from the technet script repository.  
       The script takes two parameters: a file containing a list of scripts to download, a destination directory to  
       store the downloads. If no parameters are specified the script will act on it's default values and create a file 
       containing the list of scripts to download. 
 
       This is the current list of scripts that are downloaded: 
       Search Management 
         RemoveAll-SPEnterpriseSearchCrawlRules.ps1 
       Deployment and Upgrade 
         Set-FarmProperties.ps1 
       Monitoring and Reporting 
         Report-ContentStatus.ps1 
         
.EXAMPLE  
.\Get-SharePoint2013PublicScripts.ps1 
  
 Downloads the default scripts to the current directory 
 
.EXAMPLE 
Create your own list of files to download and pass the text file as a parameter
.\Get-SharePoint2013PublicScripts.ps1 -crawlurlfile .\Get-SharePoint2013PublicScripts.txt -downloaddir C:\Temp 
 
Downloads the scripts specified in a file to a specific directory 
 
  
.LINK  
This Script - https://gallery.technet.microsoft.com/SharePoint-2013-a725a93b
.NOTES  
  File Name : Get-SharePoint2013PublicScripts.ps1  
  Author    : Brent Groom, Thomas Svensen   
#>  
 
param([string]$crawlurlfile="", [string]$downloaddir="$($pwd.drive.name):\SharePoint2013PublicScripts") 
 
$configurationname = $myinvocation.mycommand.name.Substring(0, $myinvocation.mycommand.name.IndexOf('.'))  
 
$tempdir = $Env:temp  
$tempfile = "$tempdir\Get-SharePoint2013PublicScriptsTempfile.htm" 
 
Function createScriptListFile()  
{   
  
$defaultfile = @"  
# Get scripts for the root directory
#\scripts\TechNet Script Repository\SharePoint\Search Management
#\scripts\TechNet Script Repository\SharePoint\Deployment and Upgrade
\SharePoint\Monitoring and Reporting
https://gallery.technet.microsoft.com/SearchBenchmarkWithSQLIO-055d68b4
https://gallery.technet.microsoft.com/Domain-Count-Extraction-e0c247e6
https://gallery.technet.microsoft.com/Builds-SP-Search-2013-10d72a25
https://gallery.technet.microsoft.com/SharePoint-2013-a725a93b
#Get-SPSearchTopologyState script extended to check the synchronization timer job
#https://gallery.technet.microsoft.com/office/Get-SPSearchTopologyState-b7452c6a # not sure why this one isn't working
https://gallery.technet.microsoft.com/Report-CrawlFreshness-e322b3f4

"@ | Out-File "$configurationname.txt"

$global:urlfile = "$configurationname.txt" 
"Generated file $configurationname.txt" 
}  
     
$debug=$false
 
function main
{ 
    if($global:urlfile.length -eq 0) 
    { 
        "Creating default download file to retrieve scripts" 
        createScriptListFile 
    } 
 
    $scriptpath = $downloaddir 
         
    # Iterate all lines in the file 
	Get-Content $global:urlfile |% {  
	 
	$_ = $_.Trim()
	# Create a directory based on comment line and put all files into this dir until the next comment line 
	if ($_.StartsWith('\') ) 
	{ 
		$_ = $_.substring(1) 
		new-item  -path $downloaddir -name $_ -type directory -force  | out-null 
		$scriptpath = "$downloaddir\$_" 
	} 
	# Skip blank lines and comment lines 
	elseif ($_.StartsWith('#') -OR $_.Trim().Length -eq 0 ) 
	{ 
	} 
    # Parse the technet page and get the list of files to download
    # Check the last updated timestamp from the TN page and only download if the one on disk is out of date
    else
    {
      $TNpageURL = $_
      $content = Invoke-WebRequest $TNpageURL
	  
      $downloads = $content.ParsedHtml.getElementByID("Downloads")
      $children = $downloads.childNodes
      $lastupdatedtext = $content.ParsedHtml.getElementByID("LastUpdated")
	  if($debug)
	  {
		"lastupdatedtext = $($lastupdatedtext.innerText)"
	  }
      $lastUpdatedDate = Get-Date $lastupdatedtext.innerText
      
	  foreach($child in $children)
      {
        $ScriptwebPath = $($child.pathname)
        if($ScriptwebPath -ne $null)
        {
            $webroot = $TNpageURL.substring(0,$TNpageURL.LastIndexOf('/'))
            $fullWebPath = "$webroot/$ScriptwebPath " 
            # check to see if the file already exists
            $ScriptName = $ScriptwebPath.substring($ScriptwebPath.LastIndexOf('/')+1)
			if($debug)
			{
				"Full webpath = $fullWebPath"
				"ScriptwebPath = $ScriptwebPath"
				$content.RawContent | Out-File "$ScriptName.txt"
				"ScriptName = $ScriptName"
				"Scriptpath = $scriptpath"
			}
            $fullLocalPath = "$scriptpath\$ScriptName"
            if((Test-Path -Path $fullLocalPath) -eq $false)
            {
                "Downloading $fullWebPath to $fullLocalPath"
                $wc = new-object System.Net.WebClient
                $wc.DownloadFile($fullWebPath ,"$fullLocalPath")   
				
				$theFile = Get-Item $fullLocalPath
				$theFile.CreationTime = $lastUpdatedDate
				$theFile.LastAccessTime = $lastUpdatedDate
				$theFile.LastWriteTime = $lastUpdatedDate
            }
            else
            {
				#check timestamp on the file
				$theFile = Get-Item $fullLocalPath
				$lastWriteTime = $theFile.LastWriteTime
				$ts = $lastUpdatedDate - $lastWriteTime
				# if the creation time is less than the last updated time, then download a new copy
				if($ts.TotalHours -ne 0)
				{
					$theFile = Get-Item $fullLocalPath
					$wasFileModifiedTS = $theFile.CreationTime - $theFile.LastWriteTime 
					# if the file was modified, make a backup
					if($wasFileModifiedTS.TotalMinutes -ne 0)
					{
						"Was the file modified? Creating a backup of: $fullLocalPath"					
						copy $fullLocalPath "$fullLocalPath.bak"
					}
					"Downloading updated file: $fullLocalPath"             
					$wc = new-object System.Net.WebClient
					$wc.DownloadFile($fullWebPath ,"$fullLocalPath")   
					
					$theFile = Get-Item $fullLocalPath
					$theFile.CreationTime = $lastUpdatedDate
					$theFile.LastAccessTime = $lastUpdatedDate
					$theFile.LastWriteTime = $lastUpdatedDate
				}
				"File is already current $fullLocalPath"             
            }
        }
      }
    }

	#elseif ($_.EndsWith('zip')  ) 
	#{ 
	#   "Downloading zip $_"
	#   (New-Object Net.WebClient).DownloadFile($_,"$downloaddir\temp.zip") 
	#   (new-object -com shell.application).namespace("$downloaddir").CopyHere((new-object -com shell.application).namespace("$downloaddir\temp.zip").Items(),16) 
	#   del $downloaddir\temp.zip 
	#}  
	
	} # end Get-Content 

	"Finished downloading files" 
     
} 
 
$global:urlfile = $crawlurlfile 
 
main