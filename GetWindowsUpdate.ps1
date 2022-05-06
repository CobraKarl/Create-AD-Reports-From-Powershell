# Added a countdown timer
Function Start-Countdown 
{ 
    Param(
        [Int32]$Seconds = 120,
        [string]$Message = "Pausing a while, time...."
    )
    ForEach ($Count in (1..$Seconds))
    {   Write-Progress -Id 1 -Activity $Message -Status "$($Seconds - $Count) seconds left" -PercentComplete (($Count / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}

#Check if PSUpdate is Installed

If (-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue )) {
    Write-Host "Module does not exist, Installing"
    Install-Module -Name PSWindowsUpdate -Force
    Set-ExecutionPolicy RemoteSigned -Force
    Import-Module -Name PSWindowsUpdate 
    Start-Countdown -Seconds 120
} 
Else {
    Write-Host "Module exists" 
}
  
#Enter the name that you what the report to have
$ReportName = "winupdate" 
$Report = Get-WindowsUpdate | Format-Table ComputerName, Size, Title

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


# Wait 10 Sec Before Continue
Start-Countdown -Seconds 10

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



