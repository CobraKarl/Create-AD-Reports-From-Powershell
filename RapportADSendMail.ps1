#This file creates a folder (if it do not exists, and in that folder it creats a report and then send the report to a chosen e-mail adress)

$ServerName = Read-Host -Prompt "Enter Name Of Server"
$AdminUserName = Read-Host -Prompt "Enter Admin User Name"
$AdminUserPassword = Read-Host -Prompt "Enter Password" -AsSecureString
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -Argumentlist $AdminUserName, $AdminUserPassword


# Check for source folder
$OutputPath = "C:\ServerInfo\Server"
if ((Test-Path $OutputPath) -eq $true) {


    $Report | Out-File $OutputPath\Report.html

}
else {
    # Create folder ServerInfo
    New-Item -ItemType Directory -Path C:\ServerInfo\Server

    $Report | Out-File $OutputPath\Report.html
   
}
# Create the report and covert it to a html file ( Here you can change what for information you want to have in the report)
$Report = Invoke-Command -ComputerName $ServerName { Get-ADComputer -Filter * -Properties * } -Credential $Cred | Select Name, Operatingsystem, LastLogondate@{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}  | ConvertTo-Html

# Get the credential
$credential = Get-Credential
$EmailFrom = Read-Host -Prompt "Enter E-mail Adress You Want To Send The Raport From (Must be a Office 365 e-post adress"
$EmailTo = Read-Host -Prompt "Enter E-mail Adress You Want To Send The Raport To"

## Define the Send-MailMessage parameters
$mailParams = @{
    SmtpServer                 = 'smtp.office365.com'
    Port                       = '587' 
    UseSSL                     = $true
    Credential                 = $credential
    From                       = $EmailFrom
    To                         = $EmailTo
    Subject                    = "COMPUTER RAPORT- $(Get-Date -Format g)" # Here you can change the subject information
    Body                       = 'Computer Raport'
    Attachment                 = "C:\ServerInfo\Server\Report.html" # Here you can change your attachment
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}

## Send the message
Send-MailMessage @mailParams
