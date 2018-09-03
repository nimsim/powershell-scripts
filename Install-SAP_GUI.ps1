#SAP Download Copy-Item -Source \\server\share\file -Destination C:\path\
[CmdletBinding(ConfirmImpact = 'Low', HelpURI = 'https://blog.nimtech.cloud/', SupportsPaging = $False,
    SupportsShouldProcess = $False, PositionalBinding = $False)]
Param (
    #SAP GUI parameters
    [Parameter()]$LocSap = "\\server\share\Setup\NwSapSetup.exe",
	#In case you want to download it to a location, default is install from location $LocSap
    [Parameter()]$TargetSAP = "$env:SystemRoot\Temp\NwSapSetup.exe",
    [Parameter()]$BaselineSAP = [System.Version]"7.40",
    [Parameter()]$TargetCopySAP = "$env:SystemRoot\Temp\NwSapSetup.exe",
    [Parameter()]$ArgumentSap = '/Silent /Package="SAPGUI_NWBC',

    #SAP SSO Parameters
    [Parameter()]$LocSapSSO = "\\server\share\SSO\SAPSSO.msi",
    [Parameter()]$TargetSAPSSO = "$env:SystemRoot\Temp\SAPSSO.msi",
    [Parameter()]$BaselineSAPSSO = [System.Version]"1.1.0.0",
    [Parameter()]$TargetCopySSO = "$env:SystemRoot\Temp\SAPSSO.msi",
    [Parameter()]$ArgumentSapSSO = '/qb- ALLUSERS=2',

    #All
    [Parameter()]$LogFile = "$env:ProgramData\nimtech\Logs\$($MyInvocation.MyCommand.Name).log",
    [Parameter()]$Rename = $false,
    [Parameter()]$VerbosePreference = "Continue"
)
Start-Transcript -Path $LogFile -Append

# Determine whether SAP is already installed
Write-Verbose -Message "Querying for installed SAP GUI version."
$SAP = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object  { $_.DisplayName -Like "SAP GUI*" }

# Installing SAP GUI: If SAP is not installed, download and install; or installed SAP less than current proceed with install
If (!($SAP) -or ($SAP.DisplayVersion -lt $BaselineSAP)) {

<#  Uncomment top and bottom if you want to use this, ONLY IF YOU HAVE A FULL EXE FILE FOR SAP (No dependencies)
    # Delete the installer if it exists, so that we don't have issues downloading
    If (Test-Path $TargetSAP) { Write-Verbose -Message "Deleting $TargetSAP"; Remove-Item -Path $TargetSAP -Force -ErrorAction Continue -Verbose }

    
    #Copy SAP Setup locally, works only if you have access to SAP-files e.g. company LAN or Auto-VPN; This should succeed, because the machine must have Internet access to receive the script from Intune
    #Will download regardless of network cost state (i.e. if network is marked as roaming, it will still download); Likely won't support proxy servers
    Write-Verbose -Message "Downloading SAP GUI from $LocSap"
    Copy-item -Path $LocSap -Destination $TargetSAP -ErrorAction Continue -ErrorVariable $ErrorCopy -Verbose
    
    Since our SAP version is an installer that has a lot of dependencies, we run it from a shared location with Start-Process beneath. 
#>

    # Install SAP GUI; wait 3 seconds to ensure finished; uncomment remove installer if installing a full EXE
	# Change Start-Process -Filepath to $TargetSAP if you use the download function above
    If (Test-Path $LocSap) {
        Write-Verbose -Message "Installing SAP GUI 7.40"; Start-Process -FilePath $LocSap -ArgumentList $ArgumentSap -Wait
        Write-Verbose -Message "Sleeping for 3 seconds."; Start-Sleep -Seconds 3
		
		#If Full Exe file is used, delete the setupfile after install.
        #Write-Verbose -Message "Deleting $TargetSAP"; Remove-Item -Path $TargetSAP -Force -ErrorAction Continue -Verbose
		
        Write-Verbose -Message "Querying for installed SAP version."
        $SAP = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object  { $_.DisplayName -Like "SAP GUI*" } | Select-Object DisplayName, DisplayVersion
        Write-Verbose -Message "Installed SAP: $($SAP.DisplayVersion)."
    } Else {
        $ErrorInstall = "SAP installer path at $LocSap not found."
    }

    # Intune shows basic deployment status in the Overview blade of the PowerShell script properties
    @($ErrorCopy, $ErrorInstall) | Write-Output
} Else {
    Write-Verbose "Skipping SAP installation. Installed version is $($SAP.Version)"
}

# Determine whether SAP SSO is already installed
Write-Verbose -Message "Querying for installed SAP SSO version."
$SAPSSO = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -Like "SAP Kerberos*" }

# If SAP SSO is not installed, copy and install; or installed SAP less than current proceed with install
If (!($SAPSSO) -or ($SAPSSO.Version -lt $BaselineSAPSSO)) {


    # Delete the installer if it exists, so that we don't have issues downloading
    If (Test-Path $TargetSAPSSO) { Write-Verbose -Message "Deleting $TargetSAPSSO"; Remove-Item -Path $TargetSAPSSO -Force -ErrorAction Continue -Verbose }

    # Copy SAP Setup locally; This should succeed, because the machine must have Internet access to receive the script from Intune
    # Will download regardless of network cost state (i.e. if network is marked as roaming, it will still download); Likely won't support proxy servers
    Write-Verbose -Message "Downloading SAP SSO from $LocSapSSO"
    Copy-item -Path $LocSapSSO -Destination $TargetSAPSSO -ErrorAction Continue -ErrorVariable $ErrorCopy -Verbose

    # Install SAP GUI; wait 3 seconds to ensure finished; remove installer
    If (Test-Path $TargetSAPSSO) {
        Write-Verbose -Message "Installing SAP SSO"; Start-Process -FilePath $TargetSAPSSO -ArgumentList $ArgumentSapSSO -Wait
        Write-Verbose -Message "Sleeping for 3 seconds."; Start-Sleep -Seconds 3
        Write-Verbose -Message "Deleting $TargetSAPSSO"; Remove-Item -Path $TargetSAPSSO -Force -ErrorAction Continue -Verbose
        Write-Verbose -Message "Querying for installed SAP SSO version."
        $SAPSSO = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object  { $_.DisplayName -Like "SAP Kerb*" } | Select-Object DisplayName, DisplayVersion
        Write-Verbose -Message "Installed SAP: $($SAPSSO.DisplayVersion)."
    } Else {
        $ErrorInstall = "SAP installer path at $TargetSAPSSO not found."
    }

    # Intune shows basic deployment status in the Overview blade of the PowerShell script properties
    @($ErrorCopy, $ErrorInstall) | Write-Output
} Else {
    Write-Verbose "Skipping SAP SSO installation. Installed version is $($SapSSO.Version)"
}

Stop-Transcript'