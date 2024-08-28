<#
.SYNOPSIS
   Sample Script for Updating Windows Store Apps without User interaction
.DESCRIPTION
   Update Windows Store Applications
.AUTHOR
   Marco Sap 
.VERSION
   0.5
.EXAMPLE

.DISCLAIMER
   This script code is provided as is with no guarantee or waranty
   concerning the usability or impact on systems and may be used,
   distributed, and modified in any way provided the parties agree
   and acknowledge that Microsoft or Microsoft Partners have neither
   accountabilty or responsibility for results produced by use of
   this script.

   Microsoft will not provide any support through any means.
#>

$logFile="MSStoreUpdate.log"
#Start logging
Start-Transcript "$($env:ProgramData)\Microsoft\IntuneManagementExtension\Logs\$logfile"

Write-Output "$(Get-Date -Format "dd/MM/yyyy hh:mm:ss") [Start]"

    $global:mainPFNs = @()
    $global:userRights = $false
    $global:canInstallForAllUsers = $false
    $global:allUsersSwitch = "-AllUsers"
    $global:SLEEP_DELAY = 1500

    Add-Type -AssemblyName System.ServiceModel
    $BindingFlags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
    [Windows.Management.Deployment.PackageManager,Windows.Management.Deployment,ContentType=WindowsRuntime] | Out-Null
    [Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallManager,Windows.ApplicationModel.Store.Preview.InstallControl,ContentType=WindowsRuntime] | Out-Null
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]


    Function Await($WinRtTask, $ResultType) {
      try {
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
      }
      catch {
        Write-Output "Couldn't look for Store updates, check connectivity to the Microsoft Store."   
        Write-Output "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Output "Exception Message: $($_.Exception.Message)"
       }
    }

    Function Invoke-StoreUpdates()
    {
      $global:ProgressPreference = 'Continue'
      $finished = $true
      Write-Output "[Info] Looking for all available Store updates"
      try
      {
        $appinstalls = Await ($appInstallManager.SearchForAllUpdatesAsync()) ([System.Collections.Generic.IReadOnlyList[Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallItem]])
        foreach($appinstall in $appinstalls)
        {
          if ($appinstall.PackageFamilyName)
          {
            try { $appstoreaction = ([Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallType]$appinstall.InstallType) } catch { $appstoreaction = "" }
            Write-Output "[Info] Requesting $($appinstall.PackageFamilyName) $appstoreaction"
            Start-Sleep -Milliseconds $SLEEP_DELAY
            $finished = $false
          }
        }
        Write-Output "[Info] Running the Store Update process"
        while (!$finished)
        {
          $finished = $true
          for ($index=0; $index -lt $appinstalls.Length; $index++)
          {
            $appUpdate = $appinstalls[$index]
            $packageFamilyName = $appUpdate.PackageFamilyName
            $status = $appUpdate.GetCurrentStatus()
            $currentstate = [Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallState]$status.InstallState

            if ($status.PercentComplete -eq 100)
            {
              Write-Progress -Id $index -Activity $packageFamilyName -Status "Completed" -Completed
            }
            elseif ($status.ErrorCode)
            {
              Write-Output "[Error] $packageFamilyName failed with Error $status.ErrorCode / $currentstate"
            }
            else
            {
              #Write-Progress -Id $index -Activity $packageFamilyName -status ("$currentstate $([Math]::Round($status.BytesDownloaded/1024).ToString('N0'))kb ($($status.PercentComplete)%)") -percentComplete $status.PercentComplete
              if ($finished)
              {
                $finished = $false
              }
            }
          }
          Start-Sleep -Milliseconds $SLEEP_DELAY
        }
        Write-Output "[Info] The Store Update process completed for $($appinstalls.Length) packages"
      }
      catch
      {
        Write-Output "Unable to update the Store application"
        Write-Output "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Output "Exception Message: $($_.Exception.Message)"
       }
    }
$appInstallManager = New-Object Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallManager
Invoke-StoreUpdates

Write-Output "$(Get-Date -Format "dd/MM/yyyy hh:mm:ss") [End]"

#Stop logging
Stop-Transcript