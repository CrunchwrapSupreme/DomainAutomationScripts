# Generates a .csv report on systems with a mis-matched NETBIOS nmae
# Created by David Young on 1/10/2019
# Requires Remote Server Administration Tools from Microsoft to run
Import-Module ActiveDirectory

$filedate = (Get-Date).ToString() -replace " ", "_"
$filedate = $filedate -replace ":", "_"
$filedate = $filedate -replace "/", "-"
$errfilepath = ".\Reports\ValidateADCSConnectionErrors$($filedate).txt"
$csvpath = ".\Reports\ComputerNameSerialNumberMismatchReport$($filedate).csv"

"Error report for ValidateADComputerSerial.ps1" | Out-File -FilePath $errfilepath
"=" * 80 | Out-File -Append -FilePath $errfilepath

$c = Get-ADComputer -Filter * -Properties Name,Ipv4Address,OperatingSystem
$c = $c | Where-Object -FilterScript { 
    ($_.DistinguishedName -match "CN=Computers") -or 
    ($_.DistinguishedName -match "OU=EmployeeDesktops") 
}

$dcom = New-CimSessionOption -Protocol Dcom
$wsman = New-CimSessionOption -Protocol Wsman
$mismatched_rigs = @()

ForEach($computer in $c) {
    $os = $computer.OperatingSystem
    $name = $computer.Name
    $opt = "Nil"
    try {
        if($os -match "^(?:.*Server.*|.*Windows.*\sNT\s*|\s*|.*2000.*)$") {
            continue
        }
        elseif($os -match "^(?:.*Windows.*XP.*|.*Windows.*Vista.*|.*Windows.*7.*)$")
        {
            $opt = "DCOM"
            $session = New-CimSession -ComputerName $name -ErrorAction Stop -SessionOption $dcom -OperationTimeoutSec 5
        }
        else
        {
            $opt = "WSMAN"
            $session = New-CimSession -ComputerName $name -ErrorAction Stop -SessionOption $wsman -OperationTimeoutSec 5
        }
        
        $sn = Get-CimInstance -ErrorAction Stop -ClassName win32_bios -CimSession $session | Select-Object SerialNumber
        $sn = $sn.SerialNumber

        if ($name -ne $sn) {
            $rig = New-Object -TypeName PSObject -Property @{ComputerName=$name;IPv4Address=$ip;SerialNumber=$sn}
            $mismatched_rigs += $rig
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

$mismatched_rigs | Export-Csv -Path $csvpath -NoTypeInformation
Write-Host "Press any key to continue"
[void][System.Console]::ReadKey($true)