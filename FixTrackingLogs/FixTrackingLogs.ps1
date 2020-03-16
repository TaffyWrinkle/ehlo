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
 
#================================================= 
# FixTrackingLogs Script by Stuart Presley 
# This script will fix message tracking logs so 
# that they can be imported from another machine 
# and reviewed with the -EndDate parameter on 
# GetMessageTrackingLogs 
# Use at your own risk. Make file level backups 
# of the tracking log files before running 
# 
# Requires: Exchange Management Shell 
# Usage: .\FixTrackingLogs.ps1 
# 
#================================================= 
 
function GetMachineName 
{ 
  (Get-WmiObject Win32_ComputerSystem).Name; 
} 
 
function ReadFile($item) 
{ 
  $file = [System.IO.File]::OpenText($item.FullName) 
 
  #Get the first date entry...ignoring the first 5 lines 
  for($count = 0; $count -le 5; $count++) 
  { 
    $line = $file.ReadLine() 
  } 
 
  $linearray = $line.Split(',') 
 
  #Get the date 
  $CreationTime = $linearray[0] 
 
  #now lets get the last date 
  while($file.EndOfStream -ne $true) 
  { 
    $line = $file.ReadLine() 
  } 
 
  $linearray = $line.Split(',') 
  $LastModifiedTime = $linearray[0] 
  $LastAccessTime = $CreationTime 
 
  #now lets fix the times 
  $file.Close() 
  $item.CreationTime = [System.DateTime]::Parse($CreationTime) 
  $item.LastWriteTime = [System.DateTime]::Parse($LastModifiedTime) 
  $item.LastAccessTime = [System.DateTime]::Parse($LastAccessTime) 
} 
 
function FixTrackingLogs 
{ 
  Write-Warning "It is highly suggested that you make a file level backup of your message tracking logs before running this script" 
  $confirm = Read-Host "Run this script at your own risk. Are you sure you wish to continue (Y/N)" 
 
  if($confirm.StartsWith("Y") -or $confirm.StartsWith("y")) 
  { 
  } 
  else 
  { 
    break 
  } 
 
  $machinename = GetMachineName 
  $server = Get-ExchangeServer -Identity $machinename 
  $MSExchangeIS = Get-Service MSExchangeIS 
  $MSExchangeTransport = Get-Service MSExchangeTransport 
 
  if($MSExchangeIS.Status -ieq "Running" -or $MSExchangeTransport.Status -ieq "Running") 
  { 
    Write-Warning "The Microsoft Exchange Information Store or Microsoft Exchange Transport Service is running" 
    Write-Warning "You must stop these services before running this script" 
    break 
  } 
 
  if($server.IsMailboxServer) 
  { 
    $MailboxServer = Get-MailboxServer -Identity $machinename  
    $MailboxMessageTrackingLogPath = $MailboxServer.MessageTrackingLogPath.PathName 
    $dir = Get-ChildItem $MailboxMessageTrackingLogPath 
    Write-Host "Updating Mailbox tracking logs at $MailboxMessageTrackingLogPath" 
    GetTrackingLogs($dir) 
  } 
 
  if($server.IsHubTransportServer -or $server.IsEdgeServer) 
  { 
    $transportServer = Get-TransportServer -Identity $machinename 
    $TransportServerTrackingLogPath =  
    $transportServer.MessageTrackingLogPath.PathName 
 
    if($server.IsMailboxServer -and $MailboxMessageTrackingLogPath -eq $TransportServerTrackingLogPath) 
    { 
      ###Don't do anything...the files have already been processed. 
      Write-Host "Mailbox and Hub Transport Role share same log path...no need to  
      update Hub." 
    } 
    else 
    { 
      Write-Host "Updating TransportServer Tracking Logs at $TransportServerTrackingLogPath" 
      $dir = Get-ChildItem $TransportServerTrackingLogPath 
      GetTrackingLogs($dir) 
    } 
  } 
 
  Write-Host "Restarting the Microsoft Exchange Transport Log Search Service" 
  restart-Service MSExchangeTransportLogSearch 
} 
 
function GetTrackingLogs($dir) 
{ 
  foreach($item in $dir) 
  { 
    #Rule out directories and only get MSGTRK files... 
 
    if($item.mode -ne "d----" -and $item.Name -ilike("MSGTRK*")) 
    { 
      #$item.CreationTime = [System.DateTime]::Now 
      #$item.LastWriteTime = [System.DateTime]::Now 
      #$item.LastAccessTime = [System.DateTime]::Now 
      ReadFile($item) 
    } 
  } 
} 
 
FixTrackingLogs 
 
#************************************************************************ 
