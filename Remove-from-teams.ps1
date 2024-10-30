


function InstallModules  
{
      Install-Module -Name Microsoft.Graph.Teams -Force
      Install-Module -Name ImportExcel -Force


}
function ImportModules
{
 Import-Module -Name Microsoft.Graph.Teams
 Import-Module -Name ImportExcel

}
function Connecttoteams
{
    Connect-MicrosoftTeams


}


     

    # Loop through each email and attempt to remove the user from the team
    function RemoveTeamtMembersFromExcel {
        param (
            [string]$teamId,
            [string]$excelFilePath,
            [string]$sheetName = "Sheet1"
        )
    
        # Import data from Excel
        Try {
            $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
            Write-Host "Imported data from Excel:" -ForegroundColor Green
            $users | Format-Table -AutoSize
        } Catch {
            Write-Host "Failed to import Excel data. Error: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
    
        foreach ($user in $users) {
            #column name is Email
            $userEmail = $user.Email
            if (-not [string]::IsNullOrWhiteSpace( $userEmail)) {
                Write-Host "Processing email:$userEmail" -ForegroundColor Yellow
                Try {
                     
                    remove-TeamUser -GroupId $teamId   -User $userEmail
                     Write-Host "Successfully Removed  $userEmail" -ForegroundColor Green
                } Catch {
                    Write-Host "Not Removed  $userEmail. Error: $_" -ForegroundColor Blue
                }
            }  
        }
       

    }


function Verify-UsersNotInTeamFromExcel {
    param (
        [string]$teamId,
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"
    )

    # Import data from Excel
    Try {
        $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }
$teamMembers = Get-TeamUser -GroupId $teamId
            foreach ($user in $users) {
                $userEmail = $user.Email
                $member = $teamMembers | Where-Object { $_.User -eq $userEmail }
                if ($null -eq $member) {
                    Write-Host "Verification successful: $userEmail is no longer in the team."
                } else {
                    Write-Host "Verification failed: $userEmail is still a member of the team."
                }
            }
          

    }

    
   #install modules
      InstallModules 
   #import modules
      ImportModules
   #connect to teams
      Connecttoteams

# Main script execution
# Prompt user to enter the team ID and Excel file path
$teamId = Read-Host -Prompt "Please enter the Team ID"
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"
  # Call the function to remove users to the team
RemoveTeamtMembersFromExcel -teamId $teamId -excelFilePath  $excelFilePath

# Pause to ensure uses are added
Start-Sleep -Seconds 20  

# Call the function to verify users removed from team
Verify-UsersNotInTeamFromExcel -teamId $teamId -excelFilePath $excelFilePath 