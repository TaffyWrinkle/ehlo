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
# ================================================ 
# 
# Example Profile with multiple Quick use Functions 
# Copy the contents of this file to $pshome\profile.ps1 
#  
# Fhelp             Quickly gets the full help for a cmdlet 
# Ehelp              Quickly gets the examples help for a cmdlet 
# Dismount-AllDatabase        Dismounts all databases on a specified server 
# Mount-Alldatabase        Mounts all databases on a specified server 
# Reset-Eventlog        Sets all event logging levels back to default 
# Set-SharedMailboxPermisions    Grant a user "Send as" and Full Mailbox Access permission on a target mailbox 
#  
# Last Modified 12.27.2007 
# Author: matbyrd@microsoft.com 
 
 
 
# Formating Variables 
# Allows a Central place for defining color variables 
 
[string]$success = "Green"        # Color for "Positive" messages 
[string]$info     = "White"        # Color for informational messages 
[string]$warning = "Yellow"        # Color for warning messages 
[string]$fail     = "Red"        # Color for error messages 
 
# Set Alias Names 
 
new-Alias set-smp Set-SharedMailboxPermisions 
 
# Function to show the full help for a given cmdlet 
# Single expected input of the cmdlet name 
# Example: "Fhelp mount-database" 
# ================================================================================ 
 
Function fhelp { 
 
    param ([string]$cmdlet) 
 
help $cmdlet -full 
 
} 
 
 
# Function to just show the examples for a given cmdlet 
# Single expected input of the cmdlet name 
# Example:  "Ehelp mount-database" 
# ================================================================================ 
 
Function ehelp { 
 
    param ([string]$cmdlet) 
 
help $cmdlet -example 
 
} 
 
 
# Function to dismount all Databases on a specified server 
# Single required input of the servers name  
# Defaults to the local machine name 
# Can specify $true to get confirm Dialog for all databases to be dismounted 
# Example:  "Dismount-alldatabase Myserver $true" 
# ================================================================================ 
 
Function Dismount-AllDatabase { 
    param ([string]$server = (hostname),[bool]$confirm = $false) 
 
write-host " " 
get-mailboxdatabase -server $server -status | where {$_.mounted -eq $true} | foreach { write-host "Dismounting: " -nonewline -foregroundcolor $warning; write-host $_.identity -foregroundcolor $info; Dismount-database $_.identity -confirm:$confirm} 
get-publicfolderdatabase -server $server -status | where {$_.mounted -eq $true} | foreach { write-host "Dismounting PF Database: " -nonewline -foregroundcolor $warning; write-host $_.identity -foregroundcolor $info; Dismount-database $_.identity -confirm:$confirm} 
 
 
} 
 
 
# Function to mount all databases on a specified server 
# Single expected input of the servers name 
# Defaults to the local machine name 
# Example:  "Mount-alldatabase MyServer" 
# ================================================================================ 
 
Function Mount-AllDatabase { 
    param ([string]$server = (hostname)) 
 
write-host " " 
get-mailboxdatabase -server $server -status | where {$_.mounted -eq $false} | foreach { write-host "Mounting Database: " -nonewline -foregroundcolor $success; write-host $_.identity -foregroundcolor $info; mount-database $_.identity} 
get-publicfolderdatabase -server $server -status | where {$_.mounted -eq $false} | foreach { write-host "Mounting PF Database: " -nonewline -foregroundcolor $success; write-host $_.identity -foregroundcolor $info; mount-database $_.identity} 
 
} 
 
 
# Resets the event log levels on a server to the default values 
# Single expected input of the servers name (Reset-eventloglevel myserver) 
# Defaults to the local machine name 
# Example:  "Reset-Eventloglevel MyServer" 
# ================================================================================ 
 
Function Reset-EventLogLevel { 
    param ([string]$server = (hostname)) 
 
 
get-eventloglevel -server $server | where {$_.level -ne "Lowest"} | set-eventloglevel -level lowest 
 
set-eventloglevel -id "$server\MSExchange ADAccess\Validation" -level Low 
set-eventloglevel -id "$server\MSExchange ADAccess\Topology" -level Low 
 
} 
 
 
# Sets the permissions needed for one user to access another users mailbox and "Send as" them 
# Example:  "set-sharedmailboxpermisions UserA ServiceAcct" 
# ================================================================================ 
 
Function Set-SharedMailboxPermisions { 
    Param ([string]$target,[string]$granted) 
 
Add-ADPermission $target -User $granted -Extendedrights "Send As" 
Add-MailboxPermission $target -AccessRights FullAccess -user $granted  
 
} 
