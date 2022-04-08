[CmdletBinding(DefaultParameterSetName = 'ByVMHost')]
Param (
  [Parameter(ValueFromPipeline = $true, ParameterSetName = 'ByVM')]
  $VM,
  [Parameter(ValueFromPipeline = $true, ParameterSetName = 'ByVMHost')]
  $VMHost,
  [Parameter(ValueFromPipeline = $true, ParameterSetName = 'ByCluster')]
  $Cluster,
  [Parameter(ValueFromPipeline = $true, ParameterSetName = 'ByDataCenter')]
  $DataCenter
)
switch ($PSCmdlet.ParameterSetName) {
  'ByVM' {
    $viObject = Get-VM $VM
  }
  'ByVMHost' {
    $viObject = Get-VMHost $VMHost | Get-VM | Where-Object PowerState -EQ 'PoweredOn'
  }
  'ByCluster' {
    $viObject = Get-Cluster $Cluster | Get-VMHost | Get-VM | Where-Object PowerState -EQ 'PoweredOn'
  }
  'ByDataCenter' {
    $viObject = Get-Datacenter $DataCenter | Get-Cluster | Get-VMHost | Get-VM | Where-Object PowerState -EQ 'PoweredOn'
  }
  'ByAll'{
    $viObject = Get-VMHost | Get-VM | Where-Object PowerState -EQ 'PoweredOn'
  }
  Default {
    Write-Error -Message $PSItem
  }
}

$vmPingInfo = @()
if ($viObject) {
  $vmPingInfo += $viObject | ForEach-Object {
    $vm = $_.Name
    $dnsName = $_.Guest.HostName
    if ($dnsName -match 'DOMAIN') {
      # for vm that has joined org domain
      switch ($PSVersionTable.PSVersion.Major) {
        # powershell version switch
        '6' {
          # test-connection ps ver 6 returns false when ping failed
          try {
            $pingTest = Test-Connection $dnsName -Count 1 -InformationAction Ignore -ErrorAction Stop
            $ipAddress = $pingTest | ForEach-Object Replies | ForEach-Object Address | ForEach-Object IPAddressToString
            $pingStatus = 'True'
            Write-Verbose "Testing connection to computer '$dnsName' succeeded: $ipAddress" -Verbose
          }
          catch {
            $pingStatus = 'False'
            Write-Warning "Testing connection to computer '$dnsName' failed: $ipAddress"
          }
        }
        '5' {
          # test-connection ver 5 doesn't return exception if ping failed, it returns false
          try { $pingTest = Test-Connection $dnsName -Count 1 -Quiet -ErrorAction Stop }
          catch { Write-Warning $PSItem }
          $ipAddress = Test-Connection $dnsName -Count 1 | ForEach-Object IPV4Address | ForEach-Object IPAddressToString
          if ($pingTest) {
            Write-Verbose "Testing connection to computer '$dnsName' succeeded: $ipAddress" -Verbose
            $pingStatus = 'True'
          }
          else {
            $pingStatus = 'False'
            Write-Warning "Testing connection to computer '$dnsName' failed: $ipAddress"
          }
        }
      }
    }
    else {
      # for vm that hasn't join domain, try ping using network adapter that doesn't match backup
      $niclist = $_.Guest.Nics
      $niclist | ForEach-Object {
        $nic = $_
        if ($nic -match 'Network') { $vNic = $nic.Device.NetworkName } # for vm using default network
        else {
          $vPort = 'DistributedVirtualPortgroup-' + $nic.Device.NetworkName
          $vNic = (Get-VDPortgroup -Id $vPort -ErrorAction 'SilentlyContinue').Name
        }
        if ($vNic -notmatch 'BKP|Backup' -and $nic.IPAddress[0] -ne '$null') {
          switch ($PSVersionTable.PSVersion.Major) {
            # powershell version switch
            '6' {
              # test-connection ver 6 returns false when ping failed
              try {
                $pingTest = Test-Connection $nic.IPAddress[0] -Count 1 -InformationAction Ignore -ErrorAction Stop
                $ipAddress = $pingTest | ForEach-Object Replies | ForEach-Object Address | ForEach-Object IPAddressToString
                $pingStatus = 'True'
                Write-Verbose "Testing connection to computer '$vm' succeeded: $ipAddress" -Verbose
              }
              catch {
                $pingStatus = 'False'
                Write-Warning "Testing connection to computer '$vm' failed: $ipAddress"
              }
            }
            '5' {
              # test-connection ver 5 doesn't return exception if ping failed, it returns false
              try { $pingTest = Test-Connection $nic.IPAddress[0] -Count 1 -Quiet -ErrorAction Stop }
              catch { Write-Warning $PSItem }
              if ($pingTest) {
                $ipAddress = $nic.IPAddress[0]
                Write-Verbose "Testing connection to computer '$vm' succeeded: $ipAddress" -Verbose
                $pingStatus = 'True'
              }
              else {
                $pingStatus = 'False'
                Write-Warning "Testing connection to computer '$vm' failed: $ipAddress"
              }
            }
          }
        }
        else {
          $ipAddress = "NULL"
          Write-Warning "Testing connection to computer '$vm' failed: NULL"
        }
      }
    }

    $Info = { } | Select-Object Name, HostName, IPAddress, ConnectionStatus, VMhost
    $Info.Name = $vm
    $Info.HostName = $dnsName
    $Info.IPAddress = $ipAddress
    $Info.ConnectionStatus = $pingStatus
    $Info.VMhost = $_.VMHost.Name
    $Info
  }
  $vmPingInfo
}
else { Write-Verbose "Virtual Machine discovered: 'NULL'" -Verbose }
