Confirm-Prerequsites

# Login to modules
Write-Host "Waiting for Azure login..."
Connect-AzureAD | Out-Null

Write-Host "Waiting for MS365 login..."
Connect-ExchangeOnline -ShowBanner:$false

Write-Host -ForegroundColor Yellow "Generating report..."

try {
    # Get a list of all enabled, non deleted users from AzureAD
    $users = Get-AzureADUser -All $true | Select-Object UserPrincipalName, ObjectId, DisplayName, ExtensionProperty
    
    # For each user
    foreach ($user in $users) {
        Write-Host "Checking user $($user.DisplayName)"

        # Check the unified audit log for any activity in the last 90 days from this user
        $mostRecentSignIn = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -ResultSize 1 -UserIds "$($user.UserPrincipalName)"
        
        # If no activity is found
        if ($null -eq $mostRecentSignIn) {
            # Get the users manager manager
            $manager = Get-AzureADUserManager -ObjectId $user.ObjectId
            
            # And add them to the report
            $report = [PSCustomObject]@{
                Name               = $user.DisplayName
                Email              = $user.UserPrincipalName
                Manager            = $manager.DisplayName
                AccountCreatedDate = $user.ExtensionProperty.createdDateTime
            }
            
            $report | Export-Csv "inactive-users.csv" -NoTypeInformation -Append
        }
    }
    
    # Confirm report is completed
    Write-Host -ForegroundColor Green "Successfully generated report to inactive-users.csv"
}
catch {
    Throw ("Ooops! " + $error[0].ErrorDetails)
}

function Confirm-Prerequsites {
    $azureInstalled = (Get-Module -ListAvailable -Name AzureADPreview)
    $exchangeInstalled = (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
    $runAsAdmin = (([Security.Principal.WindowsPrincipal] `
                [Security.Principal.WindowsIdentity]::GetCurrent() `
        ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    
    # Checks whether we need to install a module, and if we do, checks we are running as administrator, as admin privileges are needed to install the modules
    if (-Not $runAsAdmin -and (-Not $azureInstalled -or -Not $exchangeInstalled)) {
        Write-Host -ForegroundColor Red "You are missing one of the prerequisite modules to run this script. Please rerun this script as an administrator to allow these modules to install."
        Pause
        Exit
    }   
    
    # Checks whether the AzureADPreview module is installed, and if not, installs it
    if (-Not $azureInstalled) {
        Install-Module -Name AzureADPreview
    }
    
    # Checks whether the ExchangeOnlineManagement module is installed, and if not, installs it.
    if (-Not $exchangeInstalled) {
        Install-Module -Name ExchangeOnlineManagement
    }
}