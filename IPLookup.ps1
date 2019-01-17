# Finds an associated computer to a given IP address.
# Created by David Young on 1/16/2019

$ipinput = ""

do {
    $ipinput = Read-Host “Enter a valid IP address”
} while(-not [regex]::Match($ipinput,"^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$").Success)

try {
    [System.Net.Dns]::GetHostEntry($ipinput).HostName | Write-Host
} catch {
    "Could not resolve host. IP -potentially- free." | Write-Host
}

Write-Host "Press any key to continue"
[void][System.Console]::ReadKey($true)