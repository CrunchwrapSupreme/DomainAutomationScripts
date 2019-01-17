# Move all computers whose serial number matches their AD name to the new OU
# Created by David Young on 1/14/2019
# Requires Remote Server Administration Tools from Microsoft to run
Import-Module ActiveDirectory
[xml]$configf = Get-Content .\Config=.xml
$domainstr = $configf.configuration.vals.domainstr

$filedate = (Get-Date).ToString() -replace " ", "_"
$filedate = $filedate -replace ":", "_"
$filedate = $filedate -replace "/", "-"
$errfilepath = ".\Reports\MoveDesktopsToOUErorrs$($filedate).txt"

"Error report for MoveDesktopsToNewOU.ps1" | Out-File -FilePath $errfilepath
"=" * 80 | Out-File -Append -FilePath $errfilepath

$c = Get-ADComputer -SearchBase "CN=Computers,$domainstr" -Filter * -Properties Name,Ipv4Address,OperatingSystem

$dcom = New-CimSessionOption -Protocol Dcom
$wsman = New-CimSessionOption -Protocol Wsman

ForEach($computer in $c) {
    $os = $computer.OperatingSystem
    $name = $computer.Name
    $guid = $computer.ObjectGUID
    $dn = $computer.DistinguishedName
    $opt = ""
    $optparam = ""
    try {
        if($os -match "^(?:.*Server.*|.*Windows.*\sNT\s*|\s*)$") {
            continue
        }
        elseif($os -match "^(?:.*Windows.*XP.*)$") {
            # Add XP machines to XPDesktops OU and assign them to the XPDTOP security group.
            # This is so powershell could be distributed to them via SP3 support
            $opt = "DCOM"
            $optparam = $dcom
            $destination = "OU=XPDesktops,OU=EmployeeDesktops,$domainstr"
            $group = Get-ADGroup -Identity XPDTOP
            Move-ADObject -Identity $guid -TargetPath $destination
            Add-ADGroupMember -Identity $group -Members $guid
            continue
        }
        elseif($os -match "^(?:.*Windows.*Vista.*|.*Windows.*7.*)$")
        {
            $opt = "DCOM"
            $optparam = $dcom
        }
        else
        {
            $opt = "WSMAN"
            $optparam = $wsman
        }

        $destination = 'OU=EmployeeDesktops,$domainstr'
        
        $session = New-CimSession -ComputerName $name -ErrorAction Stop -SessionOption $optparam -OperationTimeoutSec 5
        
        $sn = Get-CimInstance -ErrorAction Stop -ClassName win32_bios -CimSession $session | Select-Object SerialNumber
        $sn = $sn.SerialNumber

        if ($name -eq $sn -and
            $dn -match "^(?:CN=.*,CN=Computers,$domainstr)$") {
            Move-ADObject -Identity $guid -TargetPath $destination
        } elseif($name -ne $sn) {
            $destination = "OU=SerialMismatch,OU=EmployeeDesktops,$domainstr"
            Move-ADObject -Identity $guid -TargetPath $destination
        }
        
        $session | Remove-CimSession
    } catch {
        $currdate = (Get-Date).ToString()
        $errmsg = "Line Number $($_.InvocationInfo.ScriptLineNumber) :: $opt :: $($_.Exception.Message)"
        $line = "( $currdate )[ $name ] $errmsg"
        $line | Write-Host
        $line | Out-File -Append -FilePath $errfilepath
    }
}

$eqlen = 80 / 2 - "Complete".Length - 1
$eqline = "=" * $eqlen
$fin = "`n$eqlen Complete $eqlen"
$fin | Out-File -Append -FilePath $errfilepath