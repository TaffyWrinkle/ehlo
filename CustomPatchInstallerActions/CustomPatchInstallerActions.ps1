################################################################################# 
#  
# The sample scripts are not supported under any Microsoft standard support  
# program or service. The sample scripts are provided AS IS without warranty  
# of any kind. Microsoft further disclaims all implied warranties including, without  
# limitation, any implied warranties of merchantability or of fitness for a particular  
# purpose. The entire risk arising out of the use or performance of the sample scripts  
# and documentation remains with you. In no event shall Microsoft, its authors, or  
# anyone else involved in the creation, production, or delivery of the scripts be liable  
# for any damages whatsoever (including, without limitation, damages for loss of business  
# profits, business interruption, loss of business information, or other pecuniary loss)  
# arising out of the use of or inability to use the sample scripts or documentation,  
# even if Microsoft has been advised of the possibility of such damages 
# 
################################################################################# 
 
############################################################################### 
# Copyright (c) 2010 Microsoft Corporation.  All rights reserved. 
# 
# CustomPatchInstallerActions.ps1.template 
# 
 
############################################################################### 
# Location of folder of log actions performed by the script, to help with troubleshooting 
# 
$script:logDir = "$env:SYSTEMDRIVE\ExchangeSetupLogs" 
 
# Services need to start and stop 
$script:servicesToStart = @("FSCController", "MSExchangeSA", "MSExchangeIS", "MSExchangeTransport") 
$script:servicesToStop = @("MSExchangeSA", "MSExchangeTransport", "MSExchangeIS", "FSCController") 
 
############################################################################### 
# Log( $entry ) 
#    Add an entry to the log file 
#    Append a string to a well-known text file with a time stamp 
# Params: 
#    Args[0] - Entry to write to log 
# Returns: 
#    void 
function Log 
{ 
    $entry = $Args[0] 
 
    $line = "[{0}] {1}" -F $(get-date).ToString("HH:mm:ss"), $entry 
    add-content -Path "$script:logDir\CustomPatchInstallerActions.log" -Value $line 
} 
 
############################################################################### 
# Get the image path for fscutility.exe 
# Return: 
#   Path for fscutility.exe 
# 
function GetImagePath 
{ 
    $imagePath = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Services\FSCController").ImagePath 
     
    $parentPath = Split-Path $imagePath -Parent 
     
    return Join-Path -Path $parentPath "fscutility.exe" 
} 
 
############################################################################### 
# Perform action on specified service and verify status 
# Params: 
#   $serviceName - service name to perform action 
#   $status - service status to verify after action 
# Return: 
#   $true if succeed otherwise $false 
# 
function DoActionAndVerifyStatus([String] $serviceName, [String] $status) 
{ 
    if ($status -eq "stopped") 
    { 
        Stop-Service $serviceName -Force 
    } 
    elseif ($status -eq "running") 
    { 
        Start-Service $serviceName 
    } 
 
    if ((Get-Service $serviceName).Status -ne $status) 
    { 
        Log "Failed to bring $serviceName to status: $status"  
        return $false 
    } 
    else 
    { 
        Log "Successfully bring $serviceName to status: $status" 
        return $true 
    } 
} 
 
############################################################################### 
# Disable ForeFront for Exchange 
# 
function DisableForeFront 
{ 
    Log "Entering DisableForeFront" 
 
    foreach ($service in $script:servicesToStop) 
    { 
        if ((DoActionAndVerifyStatus $service "stopped") -ne $true) 
        { 
            exit 
        } 
    } 
 
    Log "Disabling ForeFront for Exchange" 
    $cmd = GetImagePath 
    $parameter = "/disable" 
    &$cmd $parameter 
    Log "ForeFront for Exchange disabled" 
     
    Log "Leaving DisableForeFront" 
} 
 
############################################################################### 
# Enable ForeFront for Exchange 
# 
function EnableForeFront 
{ 
    Log "Entering EnableForeFront" 
     
    Log "Enabling ForeFront for Exchange" 
    $cmd = GetImagePath 
    $parameter = "/enable" 
    &$cmd $parameter 
    Log "ForeFront for Exchange enabled" 
     
    foreach ($service in $script:servicesToStart) 
    { 
        if ((DoActionAndVerifyStatus $service "running") -ne $true) 
        { 
            exit 
        } 
    } 
     
    Log "Leaving EnableForeFront" 
} 
 
############################################################################### 
# 
# PatchRollbackActions 
# Include items to run for rollback here 
# 
function PatchRollbackActions 
{ 
    Log "Entering PatchRollbackActions" 
     
    EnableForeFront 
     
    Log "Leaving PatchRollbackActions" 
} 
 
############################################################################### 
# 
# PrePatchInstallActions 
# Include items to run before the patch here 
# 
function PrePatchInstallActions 
{ 
    Log "Entering PrePatchInstallActions" 
     
    DisableForeFront 
     
    Log "Leaving PrePatchInstallActions" 
} 
 
############################################################################### 
# 
# PostPatchInstallActions 
# Include items to run after the patch here 
# 
function PostPatchInstallActions 
{ 
    Log "Entering PostPatchInstallActions" 
     
    EnableForeFront 
     
    Log "Leaving PostPatchInstallActions" 
} 
 
############################################################################### 
# 
# Main function 
# Installer will call the cript with the following options 
# 
switch ($Args[0]) 
{ 
    {$_ -ieq "PrePatchInstallActions" } 
    { 
        PrePatchInstallActions 
        break 
    } 
 
    {$_ -ieq "PostPatchInstallActions" } 
    { 
        PostPatchInstallActions 
        break     
    } 
 
    {$_ -ieq "PatchRollbackActions"} 
    { 
        PatchRollbackActions 
        break 
    } 
} 
 
Exit 0 
