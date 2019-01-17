# Forces GPO updates on all computers in the domain
# Created by David Young on 1/10/2019
# Requires Remote Server Administration Tools from Microsoft to run
# Presently only works via HTTP (wsman)
import-module ActiveDirectory

$failed_updates = @()
$filedate = (Get-Date).ToString() -replace " ", "_"
$filedate = $filedate -replace ":", "_"
$filedate = $filedate -replace "/", "-"
$errfilepath = ".\Reports\FGPOUpdateErrors$($filedate).txt"

$adlist = Get-ADComputer -Filter * -Properties Name,OperatingSystem,Ipv4Address,LastLogonDate | 
          Select-Object Name,OperatingSystem,Ipv4Address,LastLogonDate
$adlist = $adlist | Where-Object -FilterScript { 
    ($_.DistinguishedName -match "CN=Computers") -or 
    ($_.DistinguishedName -match "OU=EmployeeDesktops") 
}

ForEach($computer in $adlist) {
    try {
        $name = $computer.Name
        $sesh = New-PSSession -ErrorAction Stop -ComputerName $name
        Invoke-Command -ErrorAction Stop -Session $sesh { gpupdate /force }
        $sesh | Exit-PSSession
    } catch {
        $currdate = (Get-Date).ToString()
        $errmsg = "Line Number $($_.InvocationInfo.ScriptLineNumber) :: $($_.Exception.Message)"
        $line = "( $currdate )[ $name ] $errmsg"
        $line | Write-Host
        $line | Out-File -Append -FilePath $errfilepath
        $failed_updates += $computer
    }
}

$report = $adlist | ?{
    $_ | Add-Member NoteProperty "GPO Update Succeeded" $(-not $failed_updates.contains($_))
    $_
}

$report | Export-Csv -Path .\Reports\GPOUpdateReport$filedate.csv -NoTypeInformation
Write-Output $failed_updates