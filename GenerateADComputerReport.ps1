# Generates a .csv report on the computer accounts registered against the AD
# Created by David Young on 1/10/2019
# Requires Remote Server Administration Tools from Microsoft to run
import-module ActiveDirectory
$filedate = (Get-Date).ToString() -replace " ", "_"
$filedate = $filedate -replace ":", "_"
$filedate = $filedate -replace "/", "-"
$comps = Get-ADComputer -Filter * -Properties * | Select-Object Name,Ipv4Address,DNSHostName,LastLogonDate,OperatingSystem,DistinguishedName `
    | Sort-Object -Property { [System.Version]$_.Ipv4Address } `
    | Export-Csv -Path .\Reports\ADComputerReport$filedate.csv -NoTypeInformation