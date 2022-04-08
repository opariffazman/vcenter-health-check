<#
.SYNOPSIS
    This script automates the process of vcenter health check.
.DESCRIPTION
    COD vCenter Health Check Reporting.Currently supports:-
    - Identifying Not Responding ESX Host
    - Verifying NTP Daemon Not Running on ESX Host
    - Connection test to Virtual Machines affected
    - Identifying ESX Host in Maintenance Mode
    - Identifying ESX Host triggered alarm (red)
    - Identifying Disconnected | Orphaned | Inaccessible | Invalid VM
    - Identifying Inaccessible Datastore
    - Identifying Datastore with less than 20 percent free-space
    - Identifying Datastore triggered alarm (red)
    - Identifying Virtual Machine with outdated vmtools version

    The script also displays verbose information such as:-
    - Host Cluster, HA/DRS State, DRS automation level
    - Virtual Machine DNSName or Hostname with Production IP
    - Alarm Information for Host or Datastore
.NOTES
    File Name   : vcenter-health-check.ps1
    Author      : ariff.azman
    Version     : 1.2
.LINK

.INPUTS
    COD vCenter information including credentials
.OUTPUTS
    HTML Formatted Table as Daily vCenter Health Check Email
    Verbose logging transcript .log file
#>

# functions
switch ($PSVersionTable.PSVersion.Major) {
  # powershell version switch
  '6' {
    function Write-Exception {
      param ($ExceptionItem)
      $exc = $exceptionItem
      $time = $(Get-Date).ToString("dd-MM-yy h:mm:ss tt") + ''
      if ($exc.Exception.ErrorCategory) {
        $item = $exc.Exception.ErrorCategory | Out-String -NoNewline
        Write-Warning -Message "$time $item."
      }
      elseif ($exc.Exception) {
        $item = $exc.Exception | Out-String -NoNewline
        Write-Warning -Message "$time $item."
      }
      else {
        $item = $exc | Out-String -NoNewline
        Write-Warning -Message "$time $item."
      }
    }
  }
  '5' {
    function Write-Exception {
      param ($ExceptionItem)
      $exc = $exceptionItem
      $time = $(Get-Date).ToString("dd-MM-yy h:mm:ss tt") + ''
      if ($exc.Exception.ErrorCategory) {
        $item = $exc.Exception.ErrorCategory | Out-String
        $itemfixed = $item.Replace([environment]::NewLine , '')
        Write-Warning -Message "$time $itemfixed."
      }
      elseif ($exc.Exception) {
        $item = $exc.Exception | Out-String
        $itemfixed = $item.Replace([environment]::NewLine , '')
        Write-Warning -Message "$time $itemfixed."
      }
      else {
        $item = $exc | Out-String
        $itemfixed = $item.Replace([environment]::NewLine , '')
        Write-Warning -Message "$time $itemfixed."
      }
    }
  }
}

function Write-Info {
  param ($Message)
  $time = $(Get-Date).ToString("dd-MM-yy h:mm:ss tt") + ''
  Write-Verbose -Message "$time $Message." -Verbose
}

function Write-Header {
  Write-Output "------------------------------------------------------------------------------------------------------------------------`n"
  Write-Output "vcenter-health-check.ps1`n"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}

function ExitScript {
  $ProgressPreference = $OriginalPref
  Stop-Transcript | Out-Null
  Write-Info "Transcript stopped, output file is $currTranscriptName"
  exit
}

# Initial Variables **************************************************************************
mode 300 # ensure pretty output'
$vCenter = 'x.x.x.x'
$vcCredPath = "$PSScriptRoot\localAdmin.Cred"
$smtpCredPath = "$PSScriptRoot\vcsmtp.Cred"
$keyPath = "$PSScriptRoot\key.txt"
$standbyCodPath = "$PSScriptRoot\standby-cod.csv"
$currDate = $(Get-Date).ToString("dd-MM-yy")
$currTranscriptName = "$PSScriptRoot\logs\transcript-$currDate.log"

Write-Header

switch ($PSVersionTable.PSVersion.Major) {
  # powershell version switch
  '6' { Start-Transcript -Path $currTranscriptName -UseMinimalHeader -Append | Out-Null }
  '5' { Start-Transcript -Path $currTranscriptName -Append | Out-Null }
}

Write-Info "Transcript started, output file is $currTranscriptName"
Set-ExecutionPolicy -ExecutionPolicy 'RemoteSigned' -Scope 'CurrentUser' -Confirm:$false | Out-Null
Write-Info "Performing the operation `"Set-ExecutionPolicy`" on target `"RemoteSigned`""
Set-PowerCLIConfiguration -InvalidCertificateAction 'Ignore' -Scope 'Session' -Confirm:$false | Out-Null
Write-Info "Performing the operation `"Update PowerCLI configuration`""

# Import .csv & .cred files **************************************************************************
Write-Info "Performing the operation `"Import-(Clixml/Csv)`" on neccessary credential & .csv files"
try {
  $standbyCod = Import-Csv -Path $standbyCodPath
  $vcCredRaw = Import-Clixml -Path $vcCredPath
  $smtpCredRaw = Import-Clixml -Path $smtpCredPath
  $key = Get-Content $keyPath
}
catch {
  Write-Exception -ExceptionItem $PSItem
  ExitScript
}

# Decrypting credentials **************************************************************************
try {
  Write-Info "Decrypting Credentials"
  $vcCredPass = ConvertTo-SecureString -String $vcCredRaw.Password -Key $key
  $vcCred = New-Object System.Management.Automation.PSCredential($vcCredRaw.UserName, $vcCredPass)
  $smtpCredPass = ConvertTo-SecureString -String $smtpCredRaw.Password -Key $key
  $smtpCred = New-Object System.Management.Automation.PSCredential($smtpCredRaw.UserName, $smtpCredPass)
}
catch {
  Write-Exception -ExceptionItem $PSItem
  ExitScript
}

# Connect vCenter Starts **************************************************************************
try {
  Connect-VIServer $vCenter -Credential $vcCred -ErrorAction 'Stop' | Out-Null
  Write-Info "Establishing connection to vCenter Server suceeded: $vCenter"
}
catch {
  Write-Exception -ExceptionItem $PSItem
  ExitScript
}

# Virtual Machine Health Check Starts **************************************************************************

# VM Disconnected/Inaccessible/Invalid/Orphaned Runtime Connection State ***************************************
$vmErrorInfo = @()

try {
  Write-Info "Discovering Virtual Machine in Not Connected State"
  $vmError = Get-View -ViewType VirtualMachine -Filter @{'RunTime.ConnectionState' = '^(?!connected).*$' } | ForEach-Object Name
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($vmError) {
  $vmErrorInfo += Get-VM $vmError | ForEach-Object {
    $Info = { } | Select-Object Name, Hostname, NumCpu, MemoryGB, ConnectionState
    $Info.Name = $_.Name
    $Info.Hostname = $_.Guest.Hostname
    $Info.NumCpu = $_.NumCpu
    $Info.MemoryGB = $_.MemoryGB
    $Info.ConnectionState = $_.ExtensionData.RunTime.ConnectionState
    $Info
  }
  Write-Exception "Virtual Machine in Not Connected State discovered: $($vmError.Count)"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $vmErrorInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else {
  Write-Info "Virtual Machine discovered: 'NULL'"
}

# VM Tools Outdated/Need Upgrade Count ************************************************************************

try {
  Write-Info "Discovering Virtual Machine Tools Need Upgrade"
  $vmToolsCount = (Get-VM | Where-Object PowerState -eq  'PoweredOn' | ForEach-Object { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -eq 'guestToolsNeedUpgrade' }).Count
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($vmToolsCount) {
  $Info = { } | Select-Object Title, Count
  $Info.Title = "VM Tools Need Upgrade"
  $Info.Count = $vmToolsCount
  $vmToolsInfo = $Info
  Write-Exception "Virtual Machine Tools Need Upgrade discovered: $vmToolsCount"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $vmToolsInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else {
  Write-Info "Virtual Machine discovered: 'NULL'"
}

# VM Snapshot Details **********************************************************************************

try {
 Write-Info "Discovering Virtual Machine Snapshots"
 $vmSnapshotInfo = Get-VM | Get-Snapshot | Sort-Object -Property SizeGB -Descending | Select-Object VM, Name, Created, @{N="CapacityGB";E={[math]::round($_.SizeGB,4)}}
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($vmSnapshotInfo) {
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $vmSnapshotInfo | Sort-Object -Property CapacityGB -Descending | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}

# ESX Host Health Check Starts **************************************************************************

# ESX Host NTP Status ***********************************************************************************

$hostNtpInfo = @()

try {
  Write-Info "Verifying ESX Host Network Time Protocol Daemon is Running"
  $hostNtpFalse = Get-VMHost | Get-VMHostService | Where-Object { $_.key -eq 'ntpd' -and $_.Running -eq $false }
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($hostNtpFalse) {
  $hostNtpInfo = $hostNtpFalse | Select-Object VMHost, Label, Policy, Running
}

# ESX Host Connection State Not Responding *************************************************************

$hostNrInfo = @()

try {
  Write-Info "Discovering ESX Host in NotResponding State"
  $hostNr = Get-VMHost | Where-Object { $_.ConnectionState -eq 'NotResponding' -and $_.PowerState -ne 'Standby' }
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($hostNr) {
  $hostNrInfo += $hostNr | ForEach-Object {
    $Info = { } | Select-Object Name, ConnectionState, Parent, HAEnabled, DrsEnabled, DrsAutomationLevel
    $Info.Name = $_.Name
    $Info.ConnectionState = $_.ConnectionState
    $Info.Parent = $_.Parent
    $ClusterHost = Get-Cluster $_.Parent
    $Info.HAEnabled = $ClusterHost.HAEnabled
    $Info.DrsEnabled = $ClusterHost.DrsEnabled
    $Info.DrsAutomationLevel = $ClusterHost.DrsAutomationLevel
    $Info
  }
  Write-Exception "ESX Host in NotResponding State discovered: $($hostNr.Count)"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $hostNrInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"

  # get vm inside not responding host
  $hostNr | ForEach-Object {
    $VMHost = $_.Name
    Write-Info "Discovering PoweredOn Virtual Machine inside ESX Host: '$VMHost'"
      if(Get-VMHost $VMHost | Get-VM){
        # VM Connection Status inside Not Responding ESX Host ***********************************************
        Write-Info "Testing connection to Virtual Machine inside ESX Host: '$VMHost'"
        $vmPingInfo = & "$PSScriptRoot\etc\Test-VMConnection.ps1" -VMHost $VMHost
        Write-Output "------------------------------------------------------------------------------------------------------------------------"
        $vmPingInfo | Format-Table -AutoSize
        Write-Output "------------------------------------------------------------------------------------------------------------------------"
      }
      else {
        Write-Info "Virtual Machine discovered: 'NULL'"
      }
    }
}
else { Write-Info "ESX Host in NotResponding State discovered: 'NULL'" }

# ESX Host in Maintenance *****************************************************************

$hostMaintenanceInfo = @()

try {
  Write-Info "Discovering ESX Host in Maintenance Mode"
  $hostMaintenance = Get-VMHost | Where-Object ConnectionState -eq 'Maintenance'
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($hostMaintenance) {
  foreach ($VMHost in $hostMaintenance) {
    # building host details
    $Info = { } | Select-Object Name, ConnectionState, Parent, HAEnabled, DrsEnabled, DrsAutomationLevel
    $Info.Name = $VMHost.Name
    $Info.ConnectionState = $VMHost.ConnectionState
    $Info.Parent = $VMHost.Parent
    $ClusterHost = Get-Cluster $VMHost.Parent
    $Info.HAEnabled = $ClusterHost.HAEnabled
    $Info.DrsEnabled = $ClusterHost.DrsEnabled
    $Info.DrsAutomationLevel = $ClusterHost.DrsAutomationLevel
    $hostMaintenanceInfo += $Info
  }
  Write-Exception "ESX Host in Maintenance Mode discovered: '$($hostMaintenance.Count)'"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $hostMaintenanceInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}

# ESX Host Connection State Disconnected ***********************************************************

$hostDcInfo = @()

try {
  Write-Info "Discovering ESX Host in Disconnected Mode"
  $hostDc = Get-VMHost | Where-Object ConnectionState -eq 'Disconnected'
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($hostDc) {
  foreach ($VMHost in $hostDc) {
    # building host details
    $Info = { } | Select-Object Name, ConnectionState, Parent, HAEnabled, DrsEnabled, DrsAutomationLevel
    $Info.Name = $VMHost.Name
    $Info.ConnectionState = $VMHost.ConnectionState
    $Info.Parent = $VMHost.Parent
    $ClusterHost = Get-Cluster $VMHost.Parent
    $Info.HAEnabled = $ClusterHost.HAEnabled
    $Info.DrsEnabled = $ClusterHost.DrsEnabled
    $Info.DrsAutomationLevel = $ClusterHost.DrsAutomationLevel
    $hostDcInfo += $Info
  }
  Write-Exception "ESX Host in Disconnected Mode discovered: '$($hostDc.Count)'"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $hostDcInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else { Write-Info "ESX Host in Disconnected Mode discovered: 'NULL'" }

# ESX Host Red Alarm **************************************************************************

$alarmHostInfo = @()
$alarmCount = 0

try {
  Write-Info "Identifying ESX Host with Red Alarm Triggered"
  $hostsTriggered = Get-VMHost | Get-View | Where-Object TriggeredAlarmState -NE $null
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($hostsTriggered) {
  foreach ($VMHost in $hostsTriggered) {
    foreach ($triggered in $VMHost.TriggeredAlarmState) {
      if ($triggered.OverallStatus -like 'red') {
        $Info = { } | Select-Object Name, AlarmInfo, Parent, HAEnabled, DrsEnabled, DrsAutomationLevel
        $alarmDef = Get-View -Id $triggered.Alarm
        $Info.Parent = Get-VMHost $VMHost.Name | ForEach-Object Parent
        $Info.Name = $VMHost.Name
        $Info.AlarmInfo = $alarmDef.Info.Name
        $ClusterHost = Get-Cluster $Info.Parent
        $Info.HAEnabled = $ClusterHost.HAEnabled
        $Info.DrsEnabled = $ClusterHost.DrsEnabled
        $Info.DrsAutomationLevel = $ClusterHost.DrsAutomationLevel
        $alarmHostInfo += $Info
        $alarmCount++
      }
    }
  }
  Write-Exception -ExceptionItem "ESX Host with Red Alarm Triggered: '$alarmCount'"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $alarmHostInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else { Write-Info "ESX Host with Red Alarm Triggered: 'NULL'" }

try {
  $hostsConnected = Get-VMHost | Where-Object ConnectionState -eq 'Connected'
  Write-Info "ESX Host in Connected State: '$($hostsConnected.Count)'"
}
catch { Write-Exception -ExceptionItem $PSItem }

# Datastore Health Check Starts **************************************************************************

try {
  $dsError = Get-Datastore | Where-Object { $_.State -EQ 'Unavailable' -and $_.Name -match 'T2|T3|T4' }
  Write-Info "Identifying Datastore with State Unavailable or Inaccessible (Tier 2, 3 & 4)"
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($dsError) {
  Write-Exception "Datastore identified: '$($dsError.Count)'"
  $dsErrorInfo = $dsError | Select-Object Name, @{n = "Free (GB)"; e = { [Math]::Round($_.FreeSpaceGB, 2) } },
  @{n = "Total (GB)"; e = { [Math]::Round($_.CapacityGB, 2) } }, `
  @{n = "Percent Free (%)"; e = { [Math]::Round($_.FreeSpaceGB / $_.CapacityGB * 100, 2) } }, State | `
    Sort-Object -Property "Percent Free (%)"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $dsErrorInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else { Write-Info "No Datastore identified as Unavailable or Inaccessible" }

try {
  $dsLow = Get-Datastore | Where-Object { ($_.FreeSpaceGB / $_.CapacityGB * 100) -lt 20 }
  Write-Info "Identifying Datastore with less than 20 percent free space (Tier 2, 3 & 4)"
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($dsLow) {
  Write-Exception "Datastore identified: '$($dsLow.Count)'"
  # build ds details
  $dsLowInfo = $dsLow | Select-Object Name, @{n = "Free (GB)"; e = { [Math]::Round($_.FreeSpaceGB, 2) } },
  @{n = "Total (GB)"; e = { [Math]::Round($_.CapacityGB, 2) } }, `
  @{n = "Percent Free (%)"; e = { [Math]::Round($_.FreeSpaceGB / $_.CapacityGB * 100, 2) } } | `
    Sort-Object -Property "Percent Free (%)"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $dsLowInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else { Write-Info "No Datastore identified with less than 20 percent free space" }

$alarmDsInfo = @()
$alarmCount = 0

try {
  Write-Info "Identifying Datastore with Critical Red Alarm Triggered"
  $dsTriggered = Get-Datastore | Get-View | Where-Object TriggeredAlarmState -NE $null
}
catch { Write-Exception -ExceptionItem $PSItem }

if ($dsTriggered) {
  foreach ($ds in $dsTriggered) {
    foreach ($triggered in $ds.TriggeredAlarmState) {
      $alarmDef = Get-View -Id $triggered.Alarm
      if ($triggered.OverallStatus -like 'red' -and $alarmDef.Info.Name -ne 'Datastore usage on disk') {
        # omit datastore usage as already visible identified on $dslowInfo
        $Info = { } | Select-Object Name, AlarmInfo
        $Info.Name = $ds.Name
        $Info.AlarmInfo = $alarmDef.Info.Name
        $alarmDsInfo += $Info
        $alarmCount++
      }
    }
  }
  Write-Exception -ExceptionItem "Datastore with Critical Red Alarm Triggered: '$alarmCount'"
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
  $alarmDsInfo | Format-Table -AutoSize
  Write-Output "------------------------------------------------------------------------------------------------------------------------"
}
else { Write-Info "ESX Host with Red Alarm Triggered: 'NULL'" }

Write-Info "Performing the operation `"Disconnect VIServer`" on target `"User: $($global:DefaultVIServer.User), Server: 10.14.221.20, Port: 443`""
Disconnect-VIServer $vCenter -Confirm:$false | Out-Null

$primary = $standbyCod | Where-Object { [datetime]::ParseExact($_.Start, "dd-MM-yy", $null) -le (Get-Date) -And [datetime]::ParseExact($_.End, "dd-MM-yy", $null) -ge (Get-Date) } | ForEach-Object Primary
$secondary = $standbyCod | Where-Object { [datetime]::ParseExact($_.Start, "dd-MM-yy", $null) -le (Get-Date) -And [datetime]::ParseExact($_.End, "dd-MM-yy", $null) -ge (Get-Date) } | ForEach-Object Secondary

# HTML content for email body
$htmlStart = @"
<style>
body { font: normal 12px Calibri, sans-serif;}
table { border: solid 1px #DDEEEE; border-collapse: collapse; border-spacing: 0;font: normal 12px Calibri, sans-serif;}
th { background-color: #DDEFEF; border: solid 1px #DDEEEE; color: #336B6B; padding: 10px; text-align: left;text-shadow: 1px 1px 1px #fff;}
td { border: solid 1px #DDEEEE; color: #333; padding: 10px; text-shadow: 1px 1px 1px #fff;}
</style>
<body>
<b>Dear COD Standby of the week,</b>
<br>Attached is the Daily Health Check Transcript in <i>.log</i> file format, kindy verify.
"@

$htmlEnd = @"
</body>
"@

if ($vmErrorInfo) { $vmErrorHtml = $vmErrorInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>Virtual Machine Inaccessible</font></b><br><br>" }
else { $vmErrorHtml = "<br><br><b><font color = Green>Virtual Machine Inaccessible: NULL</font></b><br>" }

if ($vmToolsInfo) { $vmToolsHtml = $vmToolsInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =##ffbf00>Virtual Machine Tools Outdated</font></b><br><br>" }
else { $vmToolsHtml = "<br><br><b><font color = Green>Virtual Machine Tools Outdated: NULL</font></b><br>" }

if ($vmSnapshotInfo) { $vmSnapshotHtml = $vmSnapshotInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =##ffbf00>Virtual Machine Snapshot</font></b><br><br>" }
else { $vmSnapshotHtml = "<br><br><b><font color = Green>Virtual Machine Snapshot: NULL</font></b><br>" }

if ($hostNtpInfo) { $hostNtpHtml = $hostNtpInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>ESX Host NTP Not Running</font></b><br><br>" }
else { $hostNtpHtml = "<br><br><b><font color = Green>ESX Host NTP Not Running: NULL</font></b><br>" }

if ($hostNrInfo) { $hostNrHtml = $hostNrInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =##ffbf00>ESX Host Not Responding</font></b><br><br>" }
else { $hostNrHtml = "<br><br><b><font color = Green>ESX Host Not Responding: NULL</font></b><br>" }

if ($vmPingInfo) { $vmPingHtml = $vmPingInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#0040ff>Virtual Machine Connection Info</font></b><br><br>" }
else { $vmPingHtml = "<br><br><b><font color = Green>Virtual Machine Connection Info: NULL</font></b><br>" }

if ($hostMaintenanceInfo) { $hostMaintenanceHtml = $hostMaintenanceInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>ESX Host Maintenance</font></b><br><br>" }
else { $hostMaintenanceHtml = "<br><br><b><font color = Green>ESX Host Maintenance: NULL</font></b><br>" }

if ($hostDcInfo) { $hostDcHtml = $hostDcInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>ESX Host Disconnected</font></b><br><br>" }
else { $hostDcHtml = "<br><br><b><font color = Green>ESX Host Disconnected: NULL</font></b><br>" }

if ($alarmHostInfo) { $alarmHostHtml = $alarmHostInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>ESX Host Red Alarm</font></b><br><br>" }
else { $alarmHostHtml = "<br><br><b><font color = Green>ESX Host Without Red Alarm</font></b><br>" }

if ($dsErrorInfo) { $dsErrorHtml = $dsErrorInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>Datastore Inaccessible</font></b><br><br>" }
else { $dsErrorHtml = "<br><br><b><font color = Green>Datastore Inaccessible: NULL</font></b><br>" }

if ($dsLowInfo) { $dsLowHtml = $dsLowInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>Datastore Low Free Space</font></b><br><br>" }
else { $dsLowHtml = "<br><br><b><font color = Green>Datastore Low Free Space: NULL</font></b><br>" }

if ($alarmDsInfo) { $alarmDsHtml = $alarmDsInfo | ConvertTo-Html -Fragment -PreContent "<br><br><b><font color =#ff0000>Datastore Red Alarm</font></b><br><br>" }
else { $alarmDsHtml = "<br><br><b><font color = Green>Datastore Red Alarm: NULL</font></b><br>" }


$htmlContent = `
  $htmlStart + `
  $vmErrorHtml + `
  $vmToolsHtml + `
  $vmSnapshotHtml + `
  $hostNtpHtml + `
  $hostNrHtml + `
  $vmPingHtml + `
  $hostMaintenanceHtml + `
  $hostDcHtml + `
  $alarmHostHtml + `
  $dsErrorHtml + `
  $dsLowHtml + `
  $alarmDsHtml + `
  $htmlEnd

$mailInfo = @{
  SmtpServer                 = "smtp.dir"
  Port                       = 25
  UseSsl                     = $false
  Credential                 = $smtpCred
  From                       = "vcenter@org.com.my"
  To                         = @($primary,$secondary)
  Cc                         = '{team}@org.com.my'
  # To                         = 'ariff.azman@org.com.my'
  Subject                    = "Daily vCenter Health Check COD"
  Body                       = $htmlContent
  BodyAsHtml                 = $true
  Attachments                = $currTranscriptName
  DeliveryNotificationOption = "OnFailure"
}

Write-Info "Mail Message will be sent to COD Primary & Secondary Standby of the week `"$($primary.Split('@')[0]) & $($secondary.Split('@')[0])`""

Write-Info "Transcript stopped, output file is $currTranscriptName"
Stop-Transcript | Out-Null

try { Send-MailMessage @mailInfo }
catch { Write-Exception -ExceptionItem $PSItem }
