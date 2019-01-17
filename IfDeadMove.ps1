# Moves all machines whose logon date is older than the input value and cannot be pinged.
# The destination is OU=BrokenRetiredPCs
# Created by David Young on 1/15/2019
# Requires Remote Server Administration Tools from Microsoft to run
Import-Module ActiveDirectory

[xml]$configf = Get-Content .\Config.xml
$domainstr = $configf.configuration.vals.domainstr

$pccount = 0
$target = "OU=BrokenRetiredPCs,$domainstr"

$inputdate = Read-Host “Type date in mm/dd/yyyy format”
$inputdate = [DateTime]::Parse($inputdate)

$filedate = (Get-Date).ToString() -replace " ", "_"
$filedate = $filedate -replace ":", "_"
$filedate = $filedate -replace "/", "-"
$errfilepath = ".\Reports\$($MyInvocation.ScriptName)Errors$($filedate).txt"

"Error report for $($MyInvocation.ScriptName)" | Out-File -FilePath $errfilepath
"=" * 80 | Out-File -Append -FilePath $errfilepath

$workstations = Get-ADComputer -Filter * -Properties Name,Ipv4Address,LastLogonDate,OperatingSystem
$workstations = $workstations | Where-Object -FilterScript { 
    ($_.OperatingSystem -match "^(?:Windows XP Professional|Windows 7 Professional)$")
    -and (($_.DistinguishedName -match "CN=Computers") -or
          ($_.DistinguishedName -match "OU=EmployeeDesktops"))
    -and $_.LastLogonDate -lt $inputdate
}

ForEach($computer in $workstations) {
    if(-not (Test-Connection -ComputerName $computer.Name -Quiet) -and 
       -not ($computer.DistinguishedName -match $target))
    {
        Move-ADObject -Identity ($computer.ObjectGUID) -TargetPath $target
        Write-Host "Moved $($computer.Name) to $target"
        $pccount++
    }
}


$eqlen = 80 / 2 - "Complete".Length - 1
$eqline = "=" * $eqlen
$fin = "`n$eqlen Complete $eqlen"
$fin | Out-File -Append -FilePath $errfilepath

Write-Host "Moved $pccount workstation(s) into $target"
Write-Host "Press any key to continue"
[void][System.Console]::ReadKey($true)