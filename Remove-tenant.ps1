function installazuread{
    Install-Module -Name AzureAD -Force
    Install-Module -Name ImportExcel -Force
}
function importazuread{

    Import-Module AzureAD
Import-Module ImportExcel

}

function connecttoazuread {
    Connect-AzureAD
    
    }
    
    function RemoveTenantMembersFromExcel {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"    
    )

    # Import data from Excel
    Try {
        $users = Import-Excel -Path $excelFilePath
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    foreach ($user in $users) {
        # Assuming the column name is 'Email'
        $userEmail = $user.Email

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Processing email: $userEmail" -ForegroundColor Yellow
            Try {
                $azureUser = Get-AzureADUser -Filter "Mail eq '$userEmail'" -ErrorAction Stop

                if ($null -ne $azureUser) {
                    Remove-AzureADUser -ObjectId $azureUser.ObjectId
                    Write-Host "Successfully removed $userEmail" -ForegroundColor Green
                } else {
                    Write-Host "User $userEmail not found in Azure AD." -ForegroundColor Red
                }
            } Catch {
                Write-Host "Failed to remove $userEmail. Error: $_" -ForegroundColor Blue
            }
        }
    }
}

# Define the verification method
function Verify-UserInAzureAD {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"    
    )

    # Import data from Excel
    Try {
        $users = Import-Excel -Path $excelFilePath
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # Verification of removing from Azure AD
    foreach ($user in $users) {
        $userEmail = $user.Email
        $azureUser = Get-AzureADUser -Filter "Mail eq '$userEmail'" -ErrorAction Stop

        if ($null -eq $azureUser) {
            Write-Host "Verification successful: $userEmail is no longer a tenant." -ForegroundColor Green
        } else {
            Write-Host "Verification failed: $userEmail is still a tenant." -ForegroundColor Red
        }
    }
}

# Import the AzureAD module
importazuread

# Connect to Azure AD
connecttoazuread

# Main script execution
# Prompt user to enter the file path
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"

# Call the removal function
RemoveTenantMembersFromExcel -excelFilePath $excelFilePath 

# Call the verification function
Verify-UserInAzureAD -excelFilePath $excelFilePath 
