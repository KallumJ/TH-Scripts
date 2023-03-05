Write-Host -ForegroundColor Yellow "Generating report..."

try {
    # Get a list of all enabled, non deleted users from AzureAD
    $users = Get-AzureADUser -All $true | Select-Object UserPrincipalName, ObjectId, DisplayName
    
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
                Name           = $user.DisplayName
                Email          = $user.UserPrincipalName
                Manager        = $manager.DisplayName
            }
            
            $report | Export-Csv "inactive-users.csv" -NoTypeInformation -Append
        }
    }
    
    # Confirm report is completed
     Write-Host -ForegroundColor Green "Successfully generated report to inactive-users.csv"
}
catch {
    Throw ("Ooops! " + $error[0].Exception)
}

