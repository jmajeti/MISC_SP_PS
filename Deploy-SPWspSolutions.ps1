<############################################################################# 
This script can be used to iteratively deploy large no of wsp packages in 
SharePoint 2013 farm. This script will detect if solution needs to be deployed
to GAC and/or specific to web application or has solution which needs to use
-fulltrustbindeployment parameter. 
##############################################################################>

#Load SharePoint Snapin
Add-PSSnapin Microsoft.SharePoint.PowerShell

$SolutionsDirectory = Read-Host 'Please enter directory path for solutions '
$WebApp = Read-Host 'Please enter the WebApplication Url '

#Get solutions to be depolyed
Write-Output "Getting solutions to be deployed"
$Solutions = Get-ChildItem -LiteralPath $SolutionsDirectory -Include *.wsp

#verifies if directory is empty or not
if($Solutions -eq $null)
{
    Write-Output "No wsp packages found in directory"
    Exit
}

function InstallSolution($SolutionName)
{
    Try
        {
            Install-SPSolution -Identity $Solution -GACDeployment -WebApplication $WebApp -Confirm:$false -ErrorAction Stop
        }
        Catch
        {
            if($($_.Exception.Message) -like "*this solution contains no resources scoped for a Web application*")
            {
               Install-SPSolution -Identity $Solution -GACDeployment -Confirm:$false
            }
            elseif($($_.Exception.Message) -like "*specify the -FullTrustBinDeployment parameter to suppress this warning*")
            {
                Install-SPSolution -Identity $Solution -FullTrustBinDeployment -GACDeployment -WebApplication $WebApp -Confirm:$false
            }
            else
            {
                Write-Output "Exception Message: $($_.Exception.Message)"
            }
        } 
}

function WaitForTimer($wsp) { 
    $Solution = Get-SPSolution -Identity $wsp 
    if ($Solution -ne $null)  
    { 
        $Counter = 1    
 
        Write-Output "Waiting to finish solution timer job" 
        while( ($Solution.JobExists -eq $true ) -and ( $Counter -lt 60 ) )  
        {    
            Write-Output "Please wait..." 
            sleep 5 
            $Counter++    
        } 
 
        Write-Output "Finished the solution timer job"          
    } 
}

#deployes each solution iteratively
foreach($Solution in $Solutions)
{
    $Path = $Solution.FullName
    $Solution = $Solution.Name
    Write-Host "Proceeding to deploy $Solution"

    $GetSolution = Get-SPSolution $Solution -ErrorAction SilentlyContinue
    if($GetSolution -eq $null)
    {
        Add-SPSolution -LiteralPath $Path -Confirm:$false
        InstallSolution($Solution)

    }
    elseif($GetSolution.Deployed -eq $true)
    {
        if($GetSolution.DeploymentState -eq "GlobalDeployed")
        {
            Try
            {
                Uninstall-SPSolution -Identity $Solution -Confirm:$false -ErrorAction Stop
                WaitForTimer($Solution)
                Remove-SPSolution -Identity $Solution -Confirm:$false
                Add-SPSolution -LiteralPath $Path -Confirm:$false
                InstallSolution($Solution)
            }
            Catch
            {
                Write-Output "Exception Message: $($_.Exception.Message)"
            }

        }
        else
        {
            Try
            {
                Uninstall-SPSolution -Identity $Solution -Confirm:$false -AllWebApplications -ErrorAction Stop
                WaitForTimer($Solution)
                Remove-SPSolution -Identity $Solution -Confirm:$false
                Add-SPSolution -LiteralPath $Path -Confirm:$false
                InstallSolution($Solution)
            }
            Catch
            {
                Write-Output "Exception Message: $($_.Exception.Message)"
            }
        }
            
    }
    else
    {
        Remove-SPSolution -Identity $Solution -Confirm:$false
        Try
        {
            #Add and deploy solution
            Add-SPSolution -LiteralPath $Path -Confirm:$false
            InstallSolution($Solution)
        }
        Catch
        {
            Write-Output "Exception Message: $($_.Exception.Message)"
        } 
    }
    
    #gets current state of solution
    WaitForTimer($Solution)
    Write-Output ""
    Write-Output "Getting current state of $Solution :"
    Get-SPSolution $Solution | Format-List DisplayName, SolutionId, Deployed, DeployedServers, DeploymentState, LastOperationEndTime, LastOperationDetails

}

#Unload SharePoint Snapin
Remove-PSSnapin Microsoft.SharePoint.PowerShell

Write-Host ""
Write-Host "Script execution finished."

<############################################################################# 
End of Script
##############################################################################>