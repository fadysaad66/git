

function importazuread{

    Import-Module AzureAD
Import-Module ImportExcel

}

function connecttoazuread {
    Connect-AzureAD
    
    }





  function AddTenantMembersFromExcel {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"
    )

    # Import the Excel file
    Try {
        $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # Iterate through each user email
foreach ($user in $users) {
        $userEmail = $user.Email

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Sending invitation to: $userEmail" -ForegroundColor Yellow
            Try {
                # Send invitation
                $invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $userEmail `
                                                      -InviteRedirectUrl "https://myapps.microsoft.com" `
                                                      -SendInvitationMessage $true

                Write-Host "Invitation sent to $userEmail with ID: $($invitation.Id)" -ForegroundColor Green
            } Catch {
                Write-Host "Failed to invite $userEmail. Error: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "No valid email found in this row." -ForegroundColor Red
        }
    }
}



 
# Function to verify if the user was successfully added to Azure AD
function Verify-ExternalUsers {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"
    )

    # Import the Excel file
    Try {
        $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # Check each user's invitation status
    foreach ($user in $users) {
        $userEmail = $user.Email

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            # Check if the user is in Azure AD
            $externalUser = Get-AzureADUser -Filter "Mail eq '$userEmail'" -ErrorAction SilentlyContinue

            if ($externalUser) {
                Write-Host "Verification success: $userEmail is an external user in Azure AD." -ForegroundColor Green
            } else {
                Write-Host "Verific
                ation failed: $userEmail was not found in Azure AD." -ForegroundColor Red
            }
        } else {
            Write-Host "No valid email found for verification." -ForegroundColor Red
        }
    }
}



#import moduls
 importazuread  
# connect to AzureAD 
 Connect-AzureAD

# Main script execution
# Prompt user to enter the file path
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"
# Call the invite function
AddTenantMembersFromExcel -excelFilePath $excelFilePath 

# Pause to ensure invitations are processed
Start-Sleep -Seconds 10  # Adjust time as necessary

# Call the verification function
Verify-ExternalUsers -excelFilePath $excelFilePath 