#Check if PSUpdate is Installed

If (-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue)) {
  Write-Host "Module does not exist, Installing"
  Install-Module -Name PSWindowsUpdate -Force -WarningAction Ignore -Verbose
Set-ExecutionPolicy RemoteSigned -Force
Import-Module -Name PSWindowsUpdate
}
Else {
  Write-Host "Module exists" }




#Enter the name that you what the report to have
$ReportName = Read-Host -Prompt "Enter A Name Of The Report" 

# Check for source folder
$OutputPath = "C:\ServerInfo\WInUpdate"
if ((Test-Path $OutputPath) -eq $true) {


    $Report | Out-File $OutputPath\$ReportName.txt

}
else {
    # Create folder ServerInfo
    New-Item -ItemType Directory -Path C:\ServerInfo\WinUpdate

    $Report | Out-File $OutputPath\$ReportName.txt
} 

$Report = Get-WindowsUpdate | Format-Table ComputerName, Size, Title -Verbose


# Wait 10 Sec Before Continue
Start-Sleep -Seconds 10 

# Convert the fixed width left aligned file into a collection of psobjects
$data = Get-Content $OutputPath\$ReportName.txt | Where-Object { ![string]::IsNullOrWhiteSpace($_) }

$headerString = $data[0]
$headerElements = $headerString -split "\s+" | Where-Object { $_ }
$headerIndexes = $headerElements | ForEach-Object { $headerString.IndexOf($_) }

$results = $data | Select-Object -Skip 1  | ForEach-Object {
    $props = @{}
    $line = $_
    For ($indexStep = 0; $indexStep -le $headerIndexes.Count - 1; $indexStep++) {
        $value = $null            # Assume a null value 
        $valueLength = $headerIndexes[$indexStep + 1] - $headerIndexes[$indexStep]
        $valueStart = $headerIndexes[$indexStep]
        If (($valueLength -gt 0) -and (($valueStart + $valueLength) -lt $line.Length)) {
            $value = ($line.Substring($valueStart, $valueLength)).Trim()
        }
        ElseIf ($valueStart -lt $line.Length) {
            $value = ($line.Substring($valueStart)).Trim()
        }
        $props.($headerElements[$indexStep]) = $value    
    }
    [pscustomobject]$props
} 

# Build the html from the $result
$style = @"
<style>
    th{border: 1px solid black; background: #dddddd; padding: 5px}
    td{border: 1px solid black; padding: 5px}
</style>
"@

$results | Select-Object $headerElements | ConvertTo-Html -Head $style | Set-Content $OutputPath\$ReportName.html

# Remove the txt file
Remove-Item -Path $OutputPath\$ReportName.txt 
