# Powershell script to open all excel files in a directory with a password from a file
# then resave them to a directory without an open-password
# Looks like this is the developer: https://stackoverflow.com/q/42860894/10941169
# easy to copy code here: https://pastebin.com/UCqXUYHU

param(
    $encrypted_path = "C:\my documents\vba unlock test\Encrypted",
    $decrypted_Path = "C:\my documents\vba unlock test\Decrypted\",
    $processed_Path = "C:\my documents\vba unlock test\Processed\",
    $password_Path  = "C:\my documents\vba unlock test\Passwords\Passwords.txt"
)
#$ErrorActionPreference = "SilentlyContinue"
 
# Get Current EXCEL Process ID's so they are not affected by the scripts cleanup
$currentExcelProcessIDs = (Get-Process excel).Id
$startTime = Get-Date
 
Clear-Host
 
$passwords = Get-Content -Path $password_Path
$encryptedFiles = Get-ChildItem $encrypted_path
[int] $count = $encryptedFiles.count - 1
$ExcelObj = New-Object -ComObject Excel.Application
$ExcelObj.Visible = $false
$encryptedFiles | % {
    $encryptedFile  = $_
    Write-Host "Processing" $encryptedFile.name -ForegroundColor "DarkYellow"
    Write-Host "Items remaining: " $count
    if ($encryptedFile.Extension -like "*.xls*") {
        $passwords | % {
            $password = $_
            # Attempt to open encryptedFile
            $Workbook = $ExcelObj.Workbooks.Open($encryptedFile.fullname, 1, $false, 5, $password)
            $Workbook.Activate()
 
            # if password is correct save decrypted encryptedFile to $decrypted_Path
            if ($Workbook.Worksheets.count -ne 0 ) {
                $Workbook.Password = $null
                $savePath = Join-Path $decrypted_Path $encryptedFile.Name
                Write-Host "Decrypted: " $encryptedFile.Name -f "DarkGreen"
                $Workbook.SaveAs($savePath)
                # Added to keep Excel process memory utilization in check
                $ExcelObj.Workbooks.close()
                # Move original encryptedFile to $processed_Path
                Move-Item $encryptedFile.fullname -Destination $processed_Path -Force
            }
            else {
                $ExcelObj.Workbooks.Close()
            }
        }
    }
$count--
}
# Close Document and Application
$ExcelObj.Workbooks.close()
$ExcelObj.Application.Quit()
 
$endTime = Get-Date
 
Write-Host "Processing Complete!" -f "Green"
Write-Host "Time Started   : " $startTime.ToShortTimeString()
Write-Host "Time Completed : " $endTime.ToShortTimeString()
Write-Host "Total Duration : "
$startTime - $endTime
 
# Remove any stale Excel processes created by this scripts execution
Get-Process excel `
| Where { $currentExcelProcessIDs -notcontains $_.id } `
| Stop-Process