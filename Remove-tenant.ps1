#function to install  modules
function installazuread{
    Install-Module -Name AzureAD -Force
    Install-Module -Name ImportExcel -Force
}
#unction  to import the modules
function importazuread{

    Import-Module AzureAD
Import-Module ImportExcel

}
#function to connect to AzureAD
function connecttoazuread {
    Connect-AzureAD
    
    }
    #function to remove users 
    
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

            #column name is UserPrincipalName
            $userEmail = $user.UserPrincipalName

            if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
                Write-Host "Processing email: $userEmail" -ForegroundColor Yellow
                Try {
                     
                    $azureUser = Get-AzureADUser | Where-Object { $_.UserPrincipalName -eq $userEmail }
                   
                    Remove-AzureADUser -ObjectId $azureUser.ObjectId
                   

                    Write-Host "Successfully removed $userEmail" -ForegroundColor Green
                } Catch {
                    Write-Host "Failed to remove $userEmail. Error: $_" -ForegroundColor Blue
                }
            }  
        }
    }
    
    
        #steps
#run function   InstallModules -- to run microsoft.graph.teams and importexcel
#run function  ImportModules  -- to import these modules 
#run function connecttoazuread  -- to connect to AzureAD 
#Give prompt to parameter  $excelFilePath = "path value" -- chane pass as per required 
#run function RemoveTenantMembersFromExcel  -excelFilePath $excelFilePath