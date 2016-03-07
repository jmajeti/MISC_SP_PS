Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Settings 
$IndexLocation = "E:\Data\Search15Index� #Location must be empty, will be deleted during the process! 
$SearchAppPoolName = "Search App Pool" 
$SearchAppPoolAccountName = "NIH\NIBIBSPPRSRFARM" 
$SearchServerName = (Get-ChildItem env:computername).value 
$SearchServiceName = "Search service application" 
$SearchServiceProxyName = "Search service application Proxy" 
$DatabaseName = "Search15_ADminDB" 
Write-Host -ForegroundColor Yellow "Checking if Search Application Pool exists" 
$SPAppPool = Get-SPServiceApplicationPool -Identity $SearchAppPoolName -ErrorAction SilentlyContinue

if (!$SPAppPool) 
{ 
    Write-Host -ForegroundColor Green "Creating Search Application Pool" 
    $spAppPool = New-SPServiceApplicationPool -Name $SearchAppPoolName -Account $SearchAppPoolAccountName -Verbose 
}

# Start Services search service instance 
Write-host "Start Search Service instances...." 
Start-SPEnterpriseSearchServiceInstance $SearchServerName -ErrorAction SilentlyContinue 
Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $SearchServerName -ErrorAction SilentlyContinue

Write-Host -ForegroundColor Yellow "Checking if Search Service Application exists" 
$ServiceApplication = Get-SPEnterpriseSearchServiceApplication -Identity $SearchServiceName -ErrorAction SilentlyContinue

if (!$ServiceApplication) 
{ 
    Write-Host -ForegroundColor Green "Creating Search Service Application" 
    $ServiceApplication = New-SPEnterpriseSearchServiceApplication -Partitioned -Name $SearchServiceName -ApplicationPool $spAppPool.Name  
-DatabaseName $DatabaseName 
}

Write-Host -ForegroundColor Yellow "Checking if Search Service Application Proxy exists" 
$Proxy = Get-SPEnterpriseSearchServiceApplicationProxy -Identity $SearchServiceProxyName -ErrorAction SilentlyContinue

if (!$Proxy) 
{ 
    Write-Host -ForegroundColor Green "Creating Search Service Application Proxy" 
    New-SPEnterpriseSearchServiceApplicationProxy -Partitioned -Name $SearchServiceProxyName -SearchApplication $ServiceApplication 
}


$ServiceApplication.ActiveTopology 
Write-Host $ServiceApplication.ActiveTopology

# Clone the default Topology (which is empty) and create a new one and then activate it 
Write-Host "Configuring Search Component Topology...." 
$clone = $ServiceApplication.ActiveTopology.Clone() 
$SSI = Get-SPEnterpriseSearchServiceInstance -local 
New-SPEnterpriseSearchAdminComponent �SearchTopology $clone -SearchServiceInstance $SSI 
New-SPEnterpriseSearchContentProcessingComponent �SearchTopology $clone -SearchServiceInstance $SSI 
New-SPEnterpriseSearchAnalyticsProcessingComponent �SearchTopology $clone -SearchServiceInstance $SSI 
New-SPEnterpriseSearchCrawlComponent �SearchTopology $clone -SearchServiceInstance $SSI

Remove-Item -Recurse -Force -LiteralPath $IndexLocation -ErrorAction SilentlyContinue 
mkdir -Path $IndexLocation -Force

New-SPEnterpriseSearchIndexComponent �SearchTopology $clone -SearchServiceInstance $SSI -RootDirectory $IndexLocation 
New-SPEnterpriseSearchQueryProcessingComponent �SearchTopology $clone -SearchServiceInstance $SSI 
$clone.Activate()

Write-host "Your search service application $SearchServiceName is now ready"