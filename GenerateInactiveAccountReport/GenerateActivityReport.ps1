function GenerateActivityReport {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]    
        [int] $range = 90, 
        [Parameter(Mandatory=$True)]
        [ReportType] $reportType = [ReportType]::InactiveUser
    )

    Begin {
        enum ReportType {
            ActiveUser
            InactiveUser
        }
    
        function Install-Prerequsites {
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
        
        function MSLogin {
            # Login to modules
            Write-Host "Waiting for Azure login..."
            Connect-AzureAD | Out-Null
        
            Write-Host "Waiting for MS365 login..."
            Connect-ExchangeOnline -ShowBanner:$false
        }
        
        function WriteLoggedUser {
            param (
                [string] $name
            )
        
            if ($reportType -eq [ReportType]::InactiveUser) {
                Write-Host -ForegroundColor Red "$($name) is inactive"
            } else {
                Write-Host -ForegroundColor Green "$($name) is active"
            }
        }
        
        function WriteIgnoredUser {
            param (
                [string] $name
            )
        
            if ($reportType -eq [ReportType]::InactiveUser) {
                Write-Host "$($name) is active"
            } else {
                Write-Host "$($name) is inactive"
            }
        }
    }
    
    Process {
    
    
        Install-Prerequsites
    
        MSLogin
        
        Write-Host -ForegroundColor Yellow "Generating $([ReportType].GetEnumName($reportType)) report with a range of $($range) days..."
        
        $outFile = ".\$(Get-Date -Format "ddMMyyyyHHmm").csv"
        
        try {
            # Remove previous report
            if (Test-Path $outFile) {
                Remove-Item $outFile
            }
        
            # Get a list of all enabled, non deleted users from AzureAD
            $users = Get-AzureADUser -All $true | Select-Object UserPrincipalName, ObjectId, DisplayName, ExtensionProperty
            
            # For each user
            foreach ($user in $users) {
                # Check the unified audit log for any activity within the specified range
                $mostRecentSignIn = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-$range) -EndDate (Get-Date) -ResultSize 1 -UserIds "$($user.UserPrincipalName)"
                
                # Parse the created date
                $createdDate = [DateTime]::ParseExact($user.ExtensionProperty.createdDateTime, "dd/MM/yyyy HH:mm:ss", $null)
        
                $noActivityFound = $null -eq $mostRecentSignIn
                $oldAccount = ($createdDate -le (Get-Date).AddDays(-30))
        
                $shouldLogUser = $false
                if ($reportType -eq [ReportType]::InactiveUser) {
                    # If we are logging inactive users, we should log if no activity is found, and they have an account older than 30 days
                    $shouldLogUser = $noActivityFound -and $oldAccount
                } else {
                    # If we are logging active users, we should log if activity is found
                    $shouldLogUser = -not $noActivityFound
                }
        
                # If we have determined to log this user
                if ($shouldLogUser) {
                    # Get the users manager manager
                    $manager = Get-AzureADUserManager -ObjectId $user.ObjectId
                    
                    # And add them to the report
                    $report = [PSCustomObject]@{
                        Name               = $user.DisplayName
                        Email              = $user.UserPrincipalName
                        Manager            = $manager.DisplayName
                        AccountCreatedDate = $user.ExtensionProperty.createdDateTime
                        Notes              = ""
                    }
                                
                    $report | Export-Csv $outFile -NoTypeInformation -Append
        
                    WriteLoggedUser -name $user.DisplayName
                }
                else {
                    WriteIgnoredUser -name $user.DisplayName
                }
            }
            
            # Confirm report is completed
            Write-Host -ForegroundColor Green "Successfully generated report to $($outFile)"
        }
        
        catch {
            Throw ("Ooops! " + $error[0].Exception)
        }
    }
    
    End {
    
    }
    
    
    
    
    
    
}
GenerateActivityReport
