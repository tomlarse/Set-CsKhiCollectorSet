<#
---------------------------------------------------
Created by Tom-Inge Larsen
---------------------------------------------------

Set-CsKhiCollectorSet.ps1 script start, stop or get status on KHI data collector set on all applicable Lync servers in an environment

.Notes
   - This script can be run from any computer with Lync admin tools installed
   - The script will ask for credentials for the edge servers.
   - If you get an RPC_S_SERVER_UNAVAILABLE error running the script, be sure to check that "Perfomance Logs and Alerts"
     is allowed in Windows Firewall on the remote computer

    V1.0 February 2015 Initial version
    V1.1 February 2016 Fixed isedge function
.Link
   Twitter: http://www.twitter.com/ti83
   LinkedIn: http://www.linkedin.com/in/tomingelarsen
   Blog: http://blog.codesalot.com
   Current Release: V1.1
   https://github.com/tomlarse/Set-CsKhiCollectorSet
.EXAMPLE
   Set-CSKhiCollectorSet.ps1 -Start
   Description:
   Will find Lync servers in the environment where KHI normally is run, and start the performance collector
.EXAMPLE
   Set-CSKhiCollectorSet.ps1 -Stop
   Description:
   Will find Lync servers in the environment where KHI normally is run, and stop the performance collector 
.PARAMETER Start
   Starts the Collector sets 
.PARAMETER Stop
   Stops the Collector sets
.PARAMETER GetStatus
   Starts the Collector sets
.PARAMETER Serverlist
   .csv containing servers, if automatic isn't wanted. -Site and -ExcludeEdge can not be used with this.
.PARAMETER SpecifyCredentials 
   Credentials to be used instead of current user
.PARAMETER Site
   Will limit servers to given site
.PARAMETER ExcludeEdge
   Will exclude edgeservers.
#>

param([Parameter(Mandatory = $false)]
      [switch]$Start,
      [Parameter(Mandatory = $false)]
      [switch]$Stop,
      [Parameter(Mandatory = $false)]
      [string]$Serverlist = $null,
      [Parameter(Mandatory = $false)]
      [switch]$SpecifyCredentials,
      [Parameter(Mandatory = $false)]
      [string]$Site = $null,
      [Parameter(Mandatory = $false)]
      [switch]$ExcludeEdge,
      [Parameter(Mandatory = $false)]
      [switch]$GetStatus)

function isedge($fqdn) {
    $computer = Get-CsComputer $fqdn
    $pool = Get-CsPool $computer.pool
    $edge = $false
    foreach ($obj in $pool.Services) {
        if ($obj.Contains("EdgeServer:" + $computer.pool)) {$edge = $true}
    }
}

function ResetCredentials {
    if ($SpecifyCredentials) {
        $datacollectorset.SetCredentials($credential.UserName,$credential.GetNetworkCredential().ToString())
    } else {
        $datacollectorset.SetCredentials($null,$null)
    }
}

function isKHIapplicable($fqdn) {
    $computer = Get-CsComputer $fqdn
    $pool = Get-CsPool $computer.pool
    return $pool.Services.Contains("UserServer:" + $computer.pool) -or $pool.Services.Contains("EdgeServer:" + $computer.pool) -or $pool.Services.Contains("UserDatabase:" + $computer.pool) -or $pool.Services.Contains("WitnessStore:" + $computer.pool) -or $pool.Services.Contains("MonitoringDatabase:" + $computer.pool)
}

function GetLyncServers {

    $allcomputers = Get-CsComputer
    $alllyncservers = @()

    foreach ($computer in $allcomputers) {
        $pool = Get-CsPool $computer.Pool
        $servername = New-Object PSCustomObject | Select Name
        if (isKHIapplicable($computer.Fqdn)) {
            if (!(isedge($computer.Fqdn) -and $ExcludeEdge)) {
                if ($pool.Site -eq "Site:" + $site) {
                   $servername.name = $computer.fqdn
                } elseif ($site -eq "") {
                    $servername.name = $computer.fqdn
                }
                $alllyncservers += $servername
            }
        }
    }

    return $alllyncservers
}

$lyncservers = @()

if ($Serverlist -eq "") {
    $lyncservers = GetLyncServers    
}
else {
    $lyncservers = Import-Csv -Header "Name" -Path $serverlist
}

$datacollectorset = New-Object -COM Pla.DataCollectorSet

if ($SpecifyCredentials) {
        $credential = Get-Credential
    }

ResetCredentials

foreach ($server in $lyncservers) {
    $isedge = isedge($server.name)
    
    if ($isedge) {
        $edgecredential = $host.ui.PromptForCredential("Need credentials", $server.name + " is an Edge Server, please enter local admin credentials", "", "")
        $datacollectorset.SetCredentials($edgecredential.UserName,$edgecredential.GetNetworkCredential().ToString())
    }
    
    $datacollectorset.Query("KHI",$server.Name)
    $khistatus = $datacollectorset.Status
    
    if ($Start) {
        if ($khistatus -eq 0) {
            $datacollectorset.Start($false)
            $khistatus = $datacollectorset.Status
            if ($khistatus -eq 1) {
                Write-Host "KHI Collector on " $server.Name " was successfully started"
            }
        } elseif ($khistatus -eq 1) {
            Write-Host "KHI Collector on " $server.Name " was aready started"
        } else {
            Write-Host "KHI Collector on " $server.Name " is not installed correctly" 
        }
    } elseif ($stop) {
        if ($khistatus -eq 1) {
            $datacollectorset.Stop($false)
            $khistatus = $datacollectorset.Status
            if ($khistatus -eq 0) {
                Write-Host "KHI Collector on" $server.Name "was successfully stopped"
            }
        } elseif ($khistatus -eq 0) {
            Write-Host "KHI Collector on" $server.Name "was aready stopped"
        } else {
            Write-Host "KHI Collector on" $server.Name "is not installed correctly" 
        }
    } elseif ($getstatus) {
        if ($khistatus -eq 1) {
            Write-Host "KHI Collector on" $server.Name "is started"
        } elseif ($khistatus -eq 0) {
            Write-Host "KHI Collector on" $server.Name "is stopped"
        } else {
            write-host "Was not able to read KHI Collector on" $server.name 
        }
    }

    if ($isedge) {
        ResetCredentials
    }
}