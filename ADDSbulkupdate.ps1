# Define the attribute mappings for flexibility
$UserPrincipalNameAttr = "WorkEmail"
$JobTitleAttr = "JobTitle"
$DepartmentAttr = "Department"
$DisplayNameAttr = "DisplayName"
$ManagerUPNAttr = "ReportingTo"

# Define the path to the CSV file, log file, and error log file
$csvPath = "C:\Users\PATH\Documents\onpremuserupdate.csv"
$logPath = "C:\Users\PATH\Documents\user_update_log.txt"
$errorLogPath = "C:\Users\PATH\Documents\user_update_error_log.txt"

# Initialize counters
$successCount = 0
$failureCount = 0

# Import the CSV file
$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    $userSuccessful = $true  # Flag to track success for the user
    try {
        # Get the user account based on UserPrincipalName
        $userPrincipalName = $user.$UserPrincipalNameAttr
        $userAccount = Get-ADUser -Filter { UserPrincipalName -eq $userPrincipalName }

        if ($userAccount) {
            try {
                # Update user attributes except Manager
                Set-ADUser -Identity $userAccount `
                    -DisplayName $user.$DisplayNameAttr `
                    -Title $user.$JobTitleAttr `
                    -Department $user.$DepartmentAttr

                $logMessage = "Updated ${userPrincipalName} attributes (excluding Manager): $($user.$DisplayNameAttr)"
                Write-Host $logMessage
                $logMessage | Out-File -Append -FilePath $logPath

            } catch {
                $logMessage = "Error updating attributes for user: ${userPrincipalName}: $($_.Exception.Message)"
                Write-Host $logMessage -ForegroundColor Red
                $logMessage | Out-File -Append -FilePath $logPath
                $logMessage | Out-File -Append -FilePath $errorLogPath  # Save error to error log
                $userSuccessful = $false  # Mark user as failed
            }

            # Process the Manager attribute
            $managerUPN = $user.$ManagerUPNAttr
            if ($managerUPN) {
                $managerAccount = Get-ADUser -Filter { UserPrincipalName -eq $managerUPN }
                if ($managerAccount) {
                    try {
                        # Update the manager attribute
                        Set-ADUser -Identity $userAccount -Manager $managerAccount.DistinguishedName

                        $logMessage = "Updated manager for user: ${userPrincipalName}"
                        Write-Host $logMessage
                        $logMessage | Out-File -Append -FilePath $logPath

                    } catch {
                        $logMessage = "Error updating manager for user: ${userPrincipalName}: $($_.Exception.Message)"
                        Write-Host $logMessage -ForegroundColor Red
                        $logMessage | Out-File -Append -FilePath $logPath
                        $logMessage | Out-File -Append -FilePath $errorLogPath  # Save error to error log
                        $userSuccessful = $false  # Mark user as failed
                    }
                } else {
                    # Log manager not found with the template you provided
                    $logMessage = "Manager UPN $($user.$ManagerUPNAttr) not found for user $($user.$UserPrincipalNameAttr)."
                    Write-Host $logMessage -ForegroundColor Red
                    $logMessage | Out-File -Append -FilePath $logPath
                    $logMessage | Out-File -Append -FilePath $errorLogPath  # Save error to error log
                    $userSuccessful = $false  # Mark user as failed
                }
            } else {
                $logMessage = "No manager specified for user: ${userPrincipalName}"
                Write-Host $logMessage
                $logMessage | Out-File -Append -FilePath $logPath
            }

        } else {
            $logMessage = "User not found: ${userPrincipalName}"
            Write-Host $logMessage -ForegroundColor Red
            $logMessage | Out-File -Append -FilePath $logPath
            $logMessage | Out-File -Append -FilePath $errorLogPath  # Save error to error log
            $userSuccessful = $false  # Mark user as failed
        }

    } catch {
        $logMessage = "Error processing user ${userPrincipalName}: $($_.Exception.Message)"
        Write-Host $logMessage -ForegroundColor Red
        $logMessage | Out-File -Append -FilePath $logPath
        $logMessage | Out-File -Append -FilePath $errorLogPath  # Save error to error log
        $userSuccessful = $false  # Mark user as failed
    }

    # Increment counts based on success or failure
    if ($userSuccessful) {
        $successCount++
    } else {
        $failureCount++
    }
}

# Display summary
$summaryMessage = "Total Successful Updates: $successCount, Total Failed Updates: $failureCount"
Write-Host $summaryMessage
$summaryMessage | Out-File -Append -FilePath $logPath

# Display failed updates in red
if ($failureCount -gt 0) {
    Write-Host "Total Failed Updates: $failureCount" -ForegroundColor Red
} else {
    Write-Host "No failed updates."
}
