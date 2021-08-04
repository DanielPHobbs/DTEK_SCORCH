function Invoke-PSSession {

    <#
        .SYNOPSIS
        By Marc R Kellerman (mkellerman@outlook.com)
     
        Create a New-PSSession and Regsiter-PSSessionConfiguration to eliminate the double hop issue.
     
        Adaptation from code found:
        https://blogs.msdn.microsoft.com/sergey_babkins_blog/2015/03/18/another-solution-to-multi-hop-powershell-remoting/
     
        .PARAMETER ComputerName
        Array of ComputerName to create a session to.
                
        .PARAMETER Credential
        Credential for PSSession. Same credentials are used for the RunAsCredentail within the session.
     
        .PARAMETER SkipCACheck
        Advanced options for a PSSession
     
        .PARAMETER SkipCNCheck
        Advanced options for a PSSession
     
        .PARAMETER SkipRevocationCheck
        Advanced options for a PSSession
     
        .PARAMETER SkipRevocationCheck
        If PSSession is already estabilished, remove it and recreate it.
     
    #>
    
        [CmdletBinding()]
        Param(
            [parameter(Mandatory)][string[]]$ComputerName, 
            [parameter(Mandatory)][pscredential]$Credential,
            [switch]$SkipCACheck,
            [switch]$SkipCNCheck,
            [switch]$SkipRevocationCheck,
            [switch]$Unique
        )
    
        Begin   { Write-output "$(Get-Date) - $($MyInvocation.MyCommand): Begin" }
    
        Process {
    
        If ($Unique) { Get-PSSession -EA 0 | Where { $ComputerName -contains $_.ComputerName } | Remove-PSSession -Confirm:$False }
    
        $ConfigurationName = $Credential.GetNetworkCredential().Username
        Write-output "$(Get-Date) [Invoke-Command] Start"
        $PSSessionOption = New-PSSessionOption -SkipCACheck:$SkipCACheck.IsPresent -SkipCNCheck:$SkipCNCheck.IsPresent -SkipRevocationCheck:$SkipRevocationCheck.IsPresent
        [object[]]$PSSessionConfiguration = Invoke-Command -ComputerName $ComputerName -Credential $Credential -SessionOption $PSSessionOption -ScriptBlock { 
    
            [CmdletBinding()]Param()
            Write-output "[${Env:ComputerName}] Get-PSSessionConfiguration"
            $PSSessionConfiguration = Get-PSSessionConfiguration -Name $Using:ConfigurationName -EA 0 
            if ($PSSessionConfiguration) { Return $PSSessionConfiguration }
            Write-output "[${Env:ComputerName}] Register-PSSessionConfiguration"
            $PSSessionConfiguration = Register-PSSessionConfiguration -Name $Using:ConfigurationName -RunAsCredential $Using:Credential -MaximumReceivedDataSizePerCommandMB 1000 -MaximumReceivedObjectSizeMB 1000 -Force:$True -WA 0
            if ($PSSessionConfiguration) { Return $PSSessionConfiguration }
    
        } -EA 0 -Verbose
        Write-output "$(Get-Date) [Invoke-Command] End"
    
        
        if ($PSSessionConfiguration) { 
            Write-output "$(Get-Date) [New-PSSession] Start"
            New-PSSession -ComputerName ($ComputerName | Where { $PSSessionConfiguration.PSComputerName -Contains $_ }) -Credential $Credential -SessionOption $PSSessionOption -ConfigurationName $ConfigurationName -EA 1 
            Write-output "$(Get-Date) [New-PSSession] End"
        }
    
        }
    
        End     { Write-output "$(Get-Date) - $($MyInvocation.MyCommand): End" }
    
    }

    ####################################################################################
    
    $computername="dtekaz-hw01.dtek.com"
    $User = "DTEK\SVC-ORCH2016-RS"
    $File = "C:\secure\ScorchRA.txt"
    $PSCredential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $File | ConvertTo-SecureString)

    Invoke-PSSession -ComputerName $computername -Credential $PSCredential 