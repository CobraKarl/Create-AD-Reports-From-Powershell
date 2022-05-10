# Create-Reports-From-Powershell

<h3> RapportADSendMail.ps1 </h3>

This script creates a report on all computers that are connected to the domain on Name, OS and Last log in date (can be changed to the info you want or need). The script must be run from a computer connected to the domain. Then a folder is created and the report is placed in that folder. Then the report is sent in an email (NOTE: must be an office 365 email) to any recipient of choice.


<h3> BasicRapport.ps1 </h3>

This script creats a basic HTML computer report (you can change to the info that you want in your report)

<h3> GetWindowsUpdate.ps1 </h3>

First the script checks if PSWindowUpdate Module (that you need for runnin Get-WindowsUpdate) is installed, if it is not it will be installed otherwise the script continues then the script creats a HTML file of all updates that a are available on the computer that you run it on. It also creats a folder were the html file is saved.
Added a countdown timer just for fun :)

<h3> GetWindowsUpdateForAD.ps1 </h3>

Create a "avaiable windows updates report" from AD servern with two choices, one: Only Servers in the domain, two: All Computers
