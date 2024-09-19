## Set WinRM Service Port to a new random high-port value (or any other value desired - e.g. 389, 443 etc., assuming availability)
# v1.0 - comments to 1nTh35h311 ( yossis@protonmail.com )

# Require running elevated (PSv4.0+)
#Requires -RunAsAdministrator

# Ensure there is only one Listener set on the system
if ($(dir WSMan:\localhost\Listener).count -gt 1)
    {
        Write-Warning "[!] Found a non-standard configuration with more than one listener.`nBetter give some further thought with multiple listeners.`nQuiting automation."
        break
    }

# Get current winrm port
[int]$CurrentPort = (Get-Item WSMan:\localhost\Listener\*\Port).Value;
Write-Output "[x] Current WinRM Service Port is: $CurrentPort`n";

# Get local ports in use
$PortsInUse = netstat -ano | where {$_ -like "*TCP*0.0.0.0*LISTENING*"} | % {[int]$_.split(":")[1].split(" ")[0].trim()} | Sort-Object;
Write-Output "[x] Found $($PortsInUse.Count) Ports in use ->`n$PortsInUse";

# Generate a new random high-port
$RandomPort = Get-Random -Minimum 49152 -Maximum 65535;
while ($RandomPort -in $PortsInUse)
    {
        $RandomPort = Get-Random -Minimum 49152 -Maximum 65535;
    }

# Choose the new desired port (inc. suggetion of an available high-port)
[int]$NewPort = Read-Host "`nEnter New Port number for WinRM (e.g. $RandomPort is available),`nor press CTRL+C if you don't know what you're doing";

# Continue to set the new winrm Port
$IPAddresses = Get-ChildItem WSMan:\localhost\Listener\*\ListeningOn* | select -ExpandProperty Value | foreach {"$_`n"}
$Choice = Read-Host "The following IP Addresses will be set to listen to WinRM requests on Port $NewPort rather than $CurrentPort ->`n$IPAddresses`nContinue with setting the new port? (Y to continue, any other key to quit)"
if ($Choice -ne "Y")
    {
        break
    }

# Setting the new port
Set-Item WSMan:\localhost\Listener\*\Port -Value $NewPort -Confirm:$false -Force;

# Error handling
if (!$?)
    {
        Write-Warning "[!] An error occured while setting the new WinRM port:`n$($Error[0].Exception.Message)";
        break
    }

# Display the newly set port -> validate desired result 
[int]$PortAfterModify = (Get-Item WSMan:\localhost\Listener\*\Port).Value;
Write-Output "[x] Current WinRM Port: $PortAfterModify"