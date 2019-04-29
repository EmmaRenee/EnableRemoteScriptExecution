Function Set-WinRM
{
    Param (
        [Parameter(Mandatory=$false)] 
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory=$false)]
        [switch]$https
    )

    If (-not $PSScriptRoot) 
    {
        # Define this variable for PS 2.0
        $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
    }

    $psexec = "$PSScriptRoot\psexec.exe"
    
    If ($https)
    {
        $run = "$PSScriptRoot\winrm-https.bat"
    }
    Else
    {
        $run = "$PSScriptRoot\winrm.bat"
    }

    Foreach ($computer in $ComputerName) 
    {
        Write-Host "Processing $computer"
	
        If (Test-Connection -ComputerName $computer -Quiet -Count 1) 
        {		
            $comp = '\\' + $computer
            Write-Host "-- Running WinRM quickconfig"
            & $psexec $comp -h -accepteula -s -c $run
            Write-Host "--- Completed Computer: $computer"
        } 
        Else 
        {
            Write-Host "-- The system is offline"
        }
    }

    Write-Host "DONE"
}

Function Repair-WinRM
{
    Param (
        [Parameter(Mandatory=$false)] 
        [string[]]$ComputerName = $env:COMPUTERNAME
    )

    Foreach ($computer in $ComputerName) 
    {
        Write-Host "Processing $computer"
	
        If (Test-Connection -ComputerName $computer -Quiet -Count 1) 
        {
            Write-Host '-- Register PSSession Configuration'
        
            winrs /r:$computer powershell -noprofile -command {register-pssessionconfiguration microsoft.powershell -NoServiceRestart -Force}
            winrs /r:$computer powershell -noprofile -command {register-pssessionconfiguration microsoft.powershell32 -NoServiceRestart -Force}
          
            Write-Host '-- Restarting WinRM Service'
          
            $service =  Get-WmiObject -ComputerName $computer -Class win32_Service -Filter 'name="winrm"'
            $service.StopService()
            Start-Sleep -Seconds 2
            $service.StartService()
          
            Write-Host "-- Completed Computer: $computer"
        } 
        Else 
        {
            Write-Host '-- The system is offline'
        }
    }

    Write-Host 'DONE'
}

Function Set-ExecutionPolicyRegKeys 
{
    Param (
        [Parameter(Mandatory=$false)] 
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory=$false)]
        [ValidateSet('Undefined','AllSigned','RemoteSigned','Unrestricted','Bypass')]
        [string]$ExecutionPolicy = 'Bypass'
    )

    Foreach ($computer in $ComputerName)
    { 
        Write-Host $computer
    
        If (Test-Connection -ComputerName $computer -Quiet -Count 1)
        { 
            cmd /c reg add \\$computer\hklm\software\microsoft\powershell\1\shellds\microsoft.powershell /v executionpolicy /t reg_sz /d $ExecutionPolicy /f
            cmd /c reg add \\$computer\hklm\SOFTWARE\Policies\Microsoft\Windows\PowerShell /v executionpolicy /t reg_sz /d $ExecutionPolicy /f
            cmd /c reg add \\$computer\hklm\SOFTWARE\Policies\Microsoft\Windows\PowerShell /v enablescripts /t reg_dword /d 1 /f
        }
        Else
        {
            Write-Host "-- Offline"
        }
    }
}