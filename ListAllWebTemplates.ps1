$ver = $host | select version
if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

##
#This Script returns Site, template and Template ID for all webs in a given web application
##

##
#Set Script Parameters
##

$WebApplicationURL = "http://WebApplicationURL"
$LoggingDirectory = "E:\MyLoggingDirectory"

##
#Begin Script
##


##
#Load Functions
##

Function GenerateReport
{
#First we have to get all site collections in the web application
    $SiteCollections = Get-SPSite -WebApplication $WebApplicationURL -Limit All

#Next we'll have to check if the logging directory ends with a backslash or not, and build a log file name.
    If ($LoggingDirectory.Endswith("\"))
    {
        $LogFile = $LoggingDirectory + $SiteCollections[0].WebApplication.DisplayName + ".log"
    }
    else
    {
        $LogFile = $LoggingDirectory + "\" + $SiteCollections[0].WebApplication.DisplayName + ".log"
    }
    
#Log a file header describing what the file contains and when the file was generated
    "Site Inventory List - Generated " + (get-date) | Out-File $LogFile
    
#Log which web application is being reported
    "`r`nWeb Application: " + $SiteCollections[0].WebApplication.DisplayName | Out-File $LogFile -Append
    
#Loop through each site collection in the web application    
    Foreach($SiteCollection in $SiteCollections)
    {
    
#Record the Site Collection we're evaluating    
        "`r`nSite Collection: " + ($SiteCollection.URL) | Out-File $LogFile -Append
        
#Return a list of all webs in the site collection        
        $Webs = $SiteCollection.allwebs
        
#Loop through all of the webs in the site collection, and report the URL, TemplateID, and the Template Name.        
        foreach($web in $webs)
        {
            "Web URL: " + $Web.url | Out-File $LogFile -Append
            "Web Template ID: " + $Web.WebTemplateID | Out-File $LogFile -Append
            "Web Template: " + $Web.WebTemplate | Out-File $LogFile -Append        
        }
    }
}

##
#End Load Functions
#Begin Reporting On Web Templates
##

#Check to see if the logging directory exists
if(Test-Path $LoggingDirectory)
{

#If the logging directory exists, call the GenerateReport function
 GenerateReport
}
else
{

#If the logging directory does not exist, create it, and then call the GenerateReport function
Write-Host " "
Write-Host "Directory $LoggingDirectory Does Not Exist, Creating Directory"
New-Item $LoggingDirectory -type directory
Write-Host " "
Write-Host "Directory $LoggingDirectory Created Successfully"
Write-Host " "
GenerateReport
}

#Check to see if the logging directory was created
if (Test-Path $LoggingDirectory)
{
Write-Host "Web Templates Exported Successfully!"
}
else

#If the logging directory still does not exist, the log file cannot be created, advise the administrator
{
Write-Host "Web Templates Could Not Be Exported"
Write-Host "Please Try Again"
}