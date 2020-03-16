'#################################################################################
'# 
'# The sample scripts are not supported under any Microsoft standard support 
'# program or service. The sample scripts are provided AS IS without warranty 
'# of any kind. Microsoft further disclaims all implied warranties including, without 
'# limitation, any implied warranties of merchantability or of fitness for a particular 
'# purpose. The entire risk arising out of the use or performance of the sample scripts 
'# and documentation remains with you. In no event shall Microsoft, its authors, or 
'# anyone else involved in the creation, production, or delivery of the scripts be liable 
'# for any damages whatsoever (including, without limitation, damages for loss of business 
'# profits, business interruption, loss of business information, or other pecuniary loss) 
'# arising out of the use of or inability to use the sample scripts or documentation, 
'# even if Microsoft has been advised of the possibility of such damages
'#
'#################################################################################

'Process Message tracking log
Const strVersion="12.03" '(toddlutt 1/29/2008)

'Usage: cscript ProcessTrackingLog.vbs <LogFilePath> <NumFiles> <hub|edge|all> [ <mm/dd/yyyy> | today | yesterday ]
'Examples:
'1) To parse one file
'		cscript ProcessTrackingLog.vbs "D:\Program Files\Microsoft\Exchange Server\TransportRoles\Logs\MessageTracking\MSGTRK20070804-1.LOG" 1 all
'2) To parse one file in a single directory
'		cscript  \data\scripts\ProcessTrackingLog.vbs "D:\Program Files\Microsoft\Exchange Server\TransportRoles\Logs\MessageTracking" 1 all
'3) To parse all files in a single directory
'		cscript  \data\scripts\ProcessTrackingLog.vbs "D:\Program Files\Microsoft\Exchange Server\TransportRoles\Logs\MessageTracking" 0 all
'4) To parse all files in all subdirectories in a single directory on a central server share
'		cscript \data\scripts\ProcessTrackingLog.vbs \\CentralServer\transportlogs\MessageTracking 0 all
'5) To parse 3 files in each subdirectory on a central server share
'		cscript \data\scripts\ProcessTrackingLog.vbs \\CentralServer\transportlogs\MessageTracking 3 all
'6) To parse all files in each subdirectory that were logged yesterday on a central server share
'		cscript \data\scripts\ProcessTrackingLog.vbs \\CentralServer\transportlogs\MessageTracking 0 all yesterday
'7) To parse all files in each subdirectory that were logged on 10/29/2007 on a central server share
'		cscript \data\scripts\ProcessTrackingLog.vbs \\CentralServer\transportlogs\MessageTracking 0 all 10/29/2007
'
' NOTE: use of hub and edge assume "HUB" or "GWY" in directory path, otherwise specify all.

On Error Resume Next

Dim sStarted, sLogDateTime, sLogDateTimeEnd, sLastLogDateTimeEnd, iFileDuration, iTotalFileDateTimeDiffMin, iFileFailCount, iFileDsnFailureCount, sRuntimeDiff, sFileLine, iFileReceiveStoreDriverCount, iFileReceiveSmtpCount, iFileDeliverCount, iFailAndNdrGenerated, sLocalServerName, iFileSendSmtpCount, iMyNumFiles, strUniqueMessageId, iServerLatencyTotalCount, iDeliveryLatencyTotalCount, iServerLatencySLAMetCount, iDeliveryLatencySLAMetCount, iServerLatencyRecipientsNotCounted, iServerLatencyRecipientsCounted, iFileDumpsterTrueCount, iFileDeferCount, iFileTotalMsgBytes, iAllLastDeliveryCount, iPreTransferLatency, iPreTransferLatencyGtSLA, iTotalPreTransferLatency, iFileLastDeliveryCount, iFileTransferCount, iExceedsMaxMessageSize, sHeaderLine, iCurrentQueueCount, iMaxFileQueueCount, sQueueExceptionStart, iQueueExceptonDurationSec, iCurrentQueueSizeBytes, iMaxQueueSizeBytes, iReSubmitEventCount, iResubmitDeliveryCount, strFilterDate, iTotalFolderFilesParsed, iCurrentFileSeqNum, iLastFileSeqNum, iFilesParsedOutOfSequenceCount, iTotalFilesParsedOutOfSequenceCount
Dim aNdrMsgInfo()
Dim aRegExFileLine()
Dim aMsgSize(16,1)
Dim aMsgDeliveryLatency(24,1)
Dim aMsgServerLatency(24,1)
Dim aMsgMaxServerLatency(24,1)

'Assign Constants
Const strUsageInfo = "Usage: cscript ProcessTrackingLog.vbs <LogFilePath> <NumFiles> <hub|edge|all> [ <mm/dd/yyyy> | today | yesterday ]"
Const bLogError = FALSE
Const bPruneDictionary = FALSE
Const bWriteOrphanQueueEntries = FALSE
Const bWriteDsnFailureNotFoundResults = FALSE
Const ForReading = 1, ForWriting = 2, ForAppending = 8
'#Fields: date-time,client-ip,client-hostname,server-ip,server-hostname,source-context,connector-id,source,event-id,internal-message-id,message-id,recipient-address,recipient-status,total-bytes,recipient-count,related-recipient-address,reference,message-subject,sender-address,return-path,message-info
Const iEventTime = 0, iClientIP = 1, iClientName = 2, iServerIP = 3, iServername = 4, iSourceContext = 5, iConnectorId = 6, iSource=7, iEventID = 8, iInternalMsgId = 9, iMessageId=10, iRecipients=11, iRecipientStatus=12, iTotalBytes=13, iRecipCount = 14, iRelatedRecipAddr=15, iReference=16, iSubject=17, iSender=18, iReturnPath=19, iMessageInfo=20

'Assign Variables
bDebugOut = FALSE
iDeliveryLatencySLA = 90
iDeliveryLatencySLAExceptionLoggingThreshold = 3*iDeliveryLatencySLA
iServerLatencySLA = 30
iStageComponentLatencySLA = 5
iIndividualComponentLatencySLA = 2
iPreTransferLatencySLA = 30
iSkipFileLineMax = 10
iParseFileDurationSecMax = 300
iMaxMessageSizeThresholdKB = (2^16)
iAggregateQueueSizeThreshold = 500
sInternalDNSSuffix = "microsoft.com"
sResultPath = "c:\temp\MSGTRACK\Output\"
sArchivePath = sResultPath & "Archive\"
sSummaryResults = sResultPath & "MTSummaryResults.txt"
sReceiveResults = sResultPath & "MTReceiveResults.csv"
sNextHopResults = sResultPath & "MTNextHopResults.csv"
sLogErrors = sResultPath & "MTLogErrors.log"
sRunTimeLog = sResultPath & "MTRunTimeLog.log"
sLogStatistics = sResultPath & "MTLogStatistics.csv"
sExpandResults = sResultPath & "MTExpandResults.csv"
sDsnFailureResults = sResultPath & "MTDsnFailureResults.csv"
sDsnFailureNotFoundResults = sResultPath & "MTDsnFailureNotFoundResults.csv"
sDomainExpiredResults = sResultPath & "MTDomainExpiredResults.csv"
sMbxFullRecipResults = sResultPath & "MTMbxFullRecipResults.csv"
sTopSendersbySubmitResults = sResultPath & "MTTopSendersbySubmitResults.csv"
sTopSendersbyDeliverResults = sResultPath & "MTTopSendersbyDeliverResults.csv"
sTopRecipientResults = sResultPath & "MTTopRecipientResults.csv"
sDeliveryLatencyExceptionResults = sResultPath & "MTDeliveryLatencyExceptions.csv"
sTransferLatencyExceptionResults = sResultPath & "MTTransferLatencyExceptions.csv"
sMessageSizeExceptionResults = sResultPath & "MTMessageSizeExceptions.csv"
sQueueOrphanResults = sResultPath & "MTQueueOrphansRemoved.csv"
sDuplicateDeliveryResults = sResultPath & "MTDuplicateDeliveryResults.csv"
sFinalDeliveryResults = sResultPath & "MTFinalDeliveryResults.csv"
sLatencyTrackerResults = sResultPath & "MTLatencyTrackerResults.csv"
sRecipientNotFoundResults = sResultPath & "MTRecipientNotFoundResults.csv"
sEventTimeDistribution = sResultPath & "MTEventTimeDistribution.csv"
sMessageSizeDistribution = sResultPath & "MTMessageSizeDistribution.csv"
sBlockedDomainList = sArchivePath & "BlockedDomainList.csv"
sMailSubmissionDistribution = sResultPath & "MTMailSubmissionDistribution.csv"

iTotalFilesProcessOutOfSequenceCount = 0
iTotalFilesParsedOutOfSequenceCount = 0
iPreTransferLatencyGtSLA=0
iTotalPreTransferLatency = 0
iAllServerLatencyTotalCount = 0
iDeliveryLatencyCountLtZero = 0
iDeliveryLatencyCountUnknown = 0
iEnterParseFileFunctionTotal = 0
iRecipNotFoundCountTotal = 0
iValidReturnFailCount = 0
iSendSmtpRecipCount = 0
iReceiveSDRecipCount = 0
iReceiveSmtpRecipCount = 0
iSdDeliverRecipCount = 0
iExpandRecipCount = 0
iTotalFilesParsed = 0
iMsgIdCount = 0
iEnterFunction = 0
iUniqueDL = 0
iReceiveCount = 0
iDeliverCount = 0
iSendCount = 0
iReceiveSmtpCount = 0
iReceiveStoreDriverCount = 0
iEventCount = 0 
iTotalMsgBytes = 0
iErrorEvent = 0
iDuplicateDeliver = 0
iDuplicateDeliverRecipCount = 0
iDsnCount = 0
iResolveCount = 0
iTransferCount = 0
iExpandCount = 0
iFailCount = 0
iRedirectCount = 0
iBadmailCount = 0
iDeferCount = 0
iSplitError = 0
iMinLogEventPerMsgId = 10
iMaxLogEventPerMsgId = 0
iPoisonCount = 0
iEnterParseCsvLine = 0
iMaxAvgFailPerMin = 0
iMaxAvgFileDsnCount = 0
iTotalFileDateTimeDiffMin = 0
sOldestLogDate = now
sNewestLogDate = "01/01/1970 12:00:01 AM"
sBlockDomainDate = "01/01/1970 12:00:01 AM"
iFileDurationMax = 0
iFileDurationMin = 1000
iFailRecipientCount = 0
iFailInternalSenderCount = 0
iFailExternalSenderCount = 0
iValidReturnFailInternalSenderCount = 0
iValidReturnFailExternalSenderCount = 0
iMERDeliverCount = 0
iCORPMERDeliverCount = 0
iExternalNullReversePathDeliverCount = 0
iInternalNullReversePathDeliverCount = 0
iInternalPoisonCount = 0
iIMCEA_EX_Count = 0
sStartedScript = now
iMaxExpandRecipCount = 0
iFailAndNdrGenerated = 0
iExceedsMaxMessageSize = 0
iPercentServerLatencyRecipientsCounted = 0
iSubmitEventCount = 0

For x = 0 to 16 
	aMsgSize(x,0) = (2^x)*1024
	aMsgSize(x,1) = 0
Next	

For x = 0 to 24
	aMsgDeliveryLatency(x,0) = 2^x
	aMsgDeliveryLatency(x,1) = 0			                
Next

For x = 0 to 24
	aMsgServerLatency(x,0) = 2^x
	aMsgServerLatency(x,1) = 0			                
Next

For x = 0 to 24
	aMsgMaxServerLatency(x,0) = 2^x
	aMsgMaxServerLatency(x,1) = 0			                
Next

'Create objects
Set WshShell = CreateObject("WScript.Shell")
Set oReceiveClientDictionary = CreateObject("Scripting.Dictionary")
Set oNextHopServerDictionary = CreateObject("Scripting.Dictionary")
Set oExpandedDLDictionary = CreateObject("Scripting.Dictionary")
Set oFailReasonDictionary = CreateObject("Scripting.Dictionary")
Set oFailSourceDictionary = CreateObject("Scripting.Dictionary")
Set oTransferContextDictionary = CreateObject("Scripting.Dictionary")
Set oSDSubmitDomainDictionary = CreateObject("Scripting.Dictionary")
Set oDeferContextDictionary = CreateObject("Scripting.Dictionary")
Set oBadMailDictionary = CreateObject("Scripting.Dictionary")
Set oDsnContextDictionary = CreateObject("Scripting.Dictionary")
Set oIMCEADictionary = CreateObject("Scripting.Dictionary")
Set oRecipNotFoundDictionary = CreateObject("Scripting.Dictionary")
Set oMailboxFullDictionary = CreateObject("Scripting.Dictionary")
Set oPFSubjectDictionary = CreateObject("Scripting.Dictionary")
Set oNullReturnFailSourceDictionary = CreateObject("Scripting.Dictionary")
Set oNdrMsgIdDictionary = CreateObject("Scripting.Dictionary")
Set oNdrMsgIdNotFoundDictionary = CreateObject("Scripting.Dictionary")
Set oDomainExpiredDictionary = CreateObject("Scripting.Dictionary")
Set oTopSendersbySubmitDictionary = CreateObject("Scripting.Dictionary")
Set oTopSendersbyDeliverDictionary = CreateObject("Scripting.Dictionary")
Set oTopRecipientsDictionary = CreateObject("Scripting.Dictionary")
Set oServerLatencyDictionary = CreateObject("Scripting.Dictionary")
Set oDeliveryLatencyExceptionDictionary = CreateObject("Scripting.Dictionary")
Set oTransferDictionary = CreateObject("Scripting.Dictionary")
Set oMessageIdDictionary = CreateObject("Scripting.Dictionary")
Set oUnsortedFilesDictionary = CreateObject("Scripting.Dictionary")
Set oEmptyMessageidDictionary = CreateObject("Scripting.Dictionary")
Set oReSubmitDictionary = CreateObject("Scripting.Dictionary")
Set oReSubmitMessageDictionary = CreateObject("Scripting.Dictionary")
Set oLatencyTrackerDictionary = CreateObject("Scripting.Dictionary")
Set oDuplicateDeliveryDictionary = CreateObject("Scripting.Dictionary")
Set oFinalDeliveryDictionary = CreateObject("Scripting.Dictionary")
Set oLatencyTrackerComponentDictionary = CreateObject("Scripting.Dictionary")
Set oBlockedDomainListDictionary = CreateObject("Scripting.Dictionary")
Set oEventTimeDictionary = CreateObject("Scripting.Dictionary")
Set oSubmitDictionary = CreateObject("Scripting.Dictionary")

'PROCESSOR_ARCHITECTURE=AMD64
strProcessorArch = WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
If strProcessorArch = "AMD64" Then
	strProgramFilesPath = WshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
Else
	strProgramFilesPath = WshShell.ExpandEnvironmentStrings("%ProgramFiles%")
End If

Set fso = CreateObject("Scripting.FileSystemObject")

set oArgs = wscript.Arguments
If oArgs.count >= 3 Then
	sLogFilePath = UCASE(oArgs(0))			'Path to log file path
	iNumFiles = int(oArgs(1))
	strRole = oArgs(2)	
	Select Case oArgs.count
		Case 3
			strFilterDate = ""
		Case 4
			'Optional filter date specified
			strFilterDate = oArgs(3)
			If lcase(strFilterDate) = "today" Then strFilterDate = cstr(date)
			If lcase(strFilterDate) = "yesterday" Then strFilterDate = cstr(date - 1)
		Case Else
			'Must have 3 or 4 args
			Wscript.echo strUsageInfo
			Wscript.quit				
	End Select
Else
	Wscript.echo strUsageInfo
	Wscript.quit
End If

If fso.FolderExists(sResultPath) Then
	Wscript.echo "Results will be logged to " & sResultPath
Else
	Wscript.echo "Logging directory " & sResultPath & " does not exist"
	Wscript.echo "Please create the directory and re-run the script"
	Wscript.echo "Exiting Script"
	Wscript.quit
End If

If fso.FileExists(sSummaryResults) Then fso.DeleteFile(sSummaryResults)
Set oSummaryResults = fso.OpenTextFile(sSummaryResults,ForAppending,1)
If Err.number <> 0 Then 
	Wscript.echo "Unable to create file " & sSummaryResults & " (" & err.Description & ")"
	Wscript.echo "Exiting Script"
	Wscript.quit
End If

If fso.FileExists(sRunTimeLog) Then fso.DeleteFile(sRunTimeLog)
Set oRunTimeLog = fso.OpenTextFile(sRunTimeLog,ForAppending,1)
If Err.number <> 0 Then 
	Wscript.echo "Unable to create file " & sRunTimeLog
	Wscript.echo "Exiting Script"
	Wscript.quit
End If

LogText "Starting Script Version " & strVersion & " at " & now
LogText "Processor architecture: " & strProcessorArch
LogText GetProcessInfo(TRUE)
LogText vbcrlf 
LogText "Server Latency SLA Goal is " & iServerLatencySLA & " seconds"
LogText "End-to-End Delivery Latency SLA Goal is " & iDeliveryLatencySLA & " seconds"
LogText "Delivery Latency Exception Logging set at " & iDeliveryLatencySLAExceptionLoggingThreshold & " seconds"
LogText "Aggregate Queue Size Threshold set at " & iAggregateQueueSizeThreshold & " (Total Transport Mail Items)"
LogText "Max Message Size Threshold set at " & iMaxMessageSizeThresholdKB & " KB (details logged for messages exceeding thresold)"

If fso.FileExists(sLogErrors) Then fso.DeleteFile(sLogErrors)
Set oLogError = fso.OpenTextFile(sLogErrors,ForAppending,1)

If fso.FileExists(sNextHopResults) Then fso.DeleteFile(sNextHopResults)
Set oNextHopResults = fso.OpenTextFile(sNextHopResults,ForAppending,1)
oNextHopResults.writeline "ClientName,Source,Server,MsgCount,AvgMsgBytes,MsgCountServerSLAMet,PercentServerSLAMet,MsgCountDeliverSLAMet,PercentDeliverSLAMet"

If fso.FileExists(sReceiveResults) Then fso.DeleteFile(sReceiveResults)
Set oReceiveClientFileOut = fso.OpenTextFile(sReceiveResults,ForAppending,1)
oReceiveClientFileOut.writeline "ServerName,Source,Client,MsgCount,AvgBytes"

If fso.FileExists(sLogStatistics) Then fso.DeleteFile(sLogStatistics)
Set oLogStatistics = fso.OpenTextFile(sLogStatistics,ForAppending,1)
oLogStatistics.writeline "LogFilePath,ServerName,iLogFileEventCount,Events/MsgId,LogDateTimeStart,LogDateTimeEnd,LogDuration(Min),iAvgStoreDriverReceiveRatePerMin,iAvgSmtpReceiveRatePerMin,iAvgTransferRatePerMin,iAvgStoreDriverDeliverRatePerMin,iAvgSmtpSendRatePerMin,iAvgDsnFailRatePerMin,iAvgFailRatePerMin,iAvgDeferRatePerMin,iServerLatencyTotalCount,iPercentServerLatencySLAMet,iPercentServerLatencyRecipientsCounted,iDeliveryLatencyTotalCount,iPercentDeliveryLatencySLAMet,iPercentDumpster,iReSubmitCount,iResubmitDeliveryCount,iAvgMsgBytes,iMaxQueueCount,iQueueExceptonDurationSec,iMaxQueueSizeBytes,scriptStartTime,scriptRuntime"

If fso.FileExists(sDsnFailureResults) Then fso.DeleteFile(sDsnFailureResults)
Set oDsnFailureResults = fso.OpenTextFile(sDsnFailureResults,ForAppending,1)
oDsnFailureResults.writeline "Timestamp,ServerName,Reason,ReturnPath,Reference,RecipCount,Recipients"

If fso.FileExists(sDsnFailureNotFoundResults) Then fso.DeleteFile(sDsnFailureNotFoundResults)
If bWriteDsnFailureNotFoundResults Then
	Set oDsnFailureNotFoundResults = fso.OpenTextFile(sDsnFailureNotFoundResults,ForAppending,1)
	oDsnFailureNotFoundResults.writeline "Timestamp,ServerName,Reason,ReturnPath,MessageId,RecipCount,Recipients,Reference"
End If

If fso.FileExists(sMbxFullRecipResults) Then fso.DeleteFile(sMbxFullRecipResults)
Set oMbxFullRecipResults = fso.OpenTextFile(sMbxFullRecipResults,ForAppending,1)
oMbxFullRecipResults.writeline "Mailbox,DsnFailureCount"

If fso.FileExists(sDomainExpiredResults) Then fso.DeleteFile(sDomainExpiredResults)

If fso.FileExists(sTopSendersbySubmitResults) Then fso.DeleteFile(sTopSendersbySubmitResults)
Set oTopSendersbySubmitResults = fso.OpenTextFile(sTopSendersbySubmitResults,ForAppending,1)
oTopSendersbySubmitResults.writeline "MailboxServer,Sender,MessageCount"

If fso.FileExists(sTopSendersbyDeliverResults) Then fso.DeleteFile(sTopSendersbyDeliverResults)
Set oTopSendersbyDeliverResults = fso.OpenTextFile(sTopSendersbyDeliverResults,ForAppending,1)
oTopSendersbyDeliverResults.writeline "Sender,MessageCount"

If fso.FileExists(sTopRecipientResults) Then fso.DeleteFile(sTopRecipientResults)
Set oTopRecipientResults = fso.OpenTextFile(sTopRecipientResults,ForAppending,1)
oTopRecipientResults.writeline "MailboxServer,Recipient,MessageCount"

If fso.FileExists(sDeliveryLatencyExceptionResults) Then fso.DeleteFile(sDeliveryLatencyExceptionResults)
Set oDeliveryLatencyExceptionResults = fso.OpenTextFile(sDeliveryLatencyExceptionResults,ForAppending,1)
oDeliveryLatencyExceptionResults.writeline "MessageId,OriginalArrivalTime,TotalBytes,Sender,RecipCount,DeliverCount,MinDeliveryLatency,MaxDeliveryLatency,ClientName"

If fso.FileExists(sExpandResults) Then fso.DeleteFile(sExpandResults)
Set oExpandFileOut = fso.OpenTextFile(sExpandResults,ForAppending,1)
oExpandFileOut.writeline "Number,SMTP,RecipCount,ExpandCount,AvgExpansionLatency"

If fso.FileExists(sTransferLatencyExceptionResults) Then fso.DeleteFile(sTransferLatencyExceptionResults)
'Set oTransferLatencyExceptionResults = fso.OpenTextFile(sTransferLatencyExceptionResults,ForAppending,1)
'oTransferLatencyExceptionResults.writeline "ReceiveTime,TransferTime,MessageId,SourceContext,Latency"

If fso.FileExists(sMessageSizeExceptionResults) Then fso.DeleteFile(sMessageSizeExceptionResults)
Set oMessageSizeExceptionResults = fso.OpenTextFile(sMessageSizeExceptionResults,ForAppending,1)

If fso.FileExists(sQueueOrphanResults) Then fso.DeleteFile(sQueueOrphanResults)
If bWriteOrphanQueueEntries Then
	Set oQueueOrphanResults = fso.OpenTextFile(sQueueOrphanResults,ForAppending,1)
	oQueueOrphanResults.writeline "InternalMessageId,InternetMessageid,ReceiveTime,RecipCount,DeliverCount,EventId,Latency,LastEventTime,ReturnPath,bDumpster,TransferCount,LastTransferTime,iTotalBytes"
End If

If fso.FileExists(sDuplicateDeliveryResults) Then fso.DeleteFile(sDuplicateDeliveryResults)
Set oDuplicateDeliveryResults = fso.OpenTextFile(sDuplicateDeliveryResults,ForAppending,1)

If fso.FileExists(sFinalDeliveryResults) Then fso.DeleteFile(sFinalDeliveryResults)
Set oFinalDeliveryResults = fso.OpenTextFile(sFinalDeliveryResults,ForAppending,1)

If fso.FileExists(sLatencyTrackerResults) then fso.DeleteFile(sLatencyTrackerResults)

If fso.FileExists(sRecipientNotFoundResults) then fso.DeleteFile(sRecipientNotFoundResults)

If fso.FileExists(sEventTimeDistribution) then fso.DeleteFile(sEventTimeDistribution)

If fso.FileExists(sMessageSizeDistribution) then fso.DeleteFile(sMessageSizeDistribution)
Set oMessageSizeDistribution = fso.OpenTextFile(sMessageSizeDistribution,ForAppending,1)

If fso.FileExists(sMailSubmissionDistribution) then fso.DeleteFile(sMailSubmissionDistribution)

If Err.number <> 0 Then 
	LogText now & vbtab & "ERROR" & vbtab & "Encountered Error 0x" & HEX(err.number) & ": " & err.Description & " before loading blocked domain list"
	Err.Clear
End If

If fso.FileExists(sBlockedDomainList) Then
	LogText vbcrlf & now & vbtab & "INFO" & vbtab & "Loading Blocked Domain List" 
	Set oBlockedDomainList = fso.OpenTextFile(sBlockedDomainList,ForReading)
	Do While oBlockedDomainList.AtEndOfStream <> True
		sBlockedDomainListLine =  oBlockedDomainList.readline
		If instr(1,sBlockedDomainListLine,",") Then
			aBlockedDomainListLine = split(sBlockedDomainListLine,",")
		Else
			aBlockedDomainListLine = Array(sBlockedDomainListLine,0,0,sBlockDomainDate,0,sBlockDomainDate)
		End If
		sDomainExpiredDictionaryKey = lcase(trim(aBlockedDomainListLine(0)))
		If oDomainExpiredDictionary.Exists(sDomainExpiredDictionaryKey) Then
			LogText now & vbtab & "ERROR" & vbtab & "Duplicate blocked domain list entry found (" & sDomainExpiredDictionaryKey & ")"
		Else
			oDomainExpiredDictionary.Add sDomainExpiredDictionaryKey,Array(aBlockedDomainListLine(1),aBlockedDomainListLine(2),aBlockedDomainListLine(3),aBlockedDomainListLine(4),aBlockedDomainListLine(5))
		End If
	Loop
	LogText now & vbtab & "INFO" & vbtab & "Done Loading Blocked Domain List Containing " & oDomainExpiredDictionary.Count & " Entries"
	oBlockedDomainList.Close
	Set oBlockedDomainList = Nothing
End If

If Err.number <> 0 Then 
	LogText now & vbtab & "ERROR" & vbtab & "Encountered Error 0x" & HEX(err.number) & ": " & err.Description & " at end of loading blocked domain list"
	Err.Clear
End If

If fso.FolderExists(sLogFilePath) Then

		If GetFolders(sLogFilePath) Then
			LogText now & vbtab & "DONE" & vbtab & "Main done with files in " & sLogFilePath & " (" & iTotalFilesParsed & " total files parsed)"
		End If

Else
  
  LogText now & vbtab & "INFO" & vbtab & "Path provided is not a valid folder"
  
  If fso.FileExists(sLogFilePath) Then

    LogText now & vbtab & "INFO" & vbtab & "Processing single file: " & sLogFilePath
  
    If instr(1,lcase(sLogFilePath),".log") AND instr(1,lcase(sLogFilePath),".idx")=0 Then
  	
  	  LogText now & vbtab & "FLDR" & vbtab & "Working with " & sLogFilePath
	  
	    If instr(1,sLogFilePath,"MSGTRKM") = 0 Then
		    'Edge/Hub Tracking Log
		    ParseFile sLogFilePath
	    Else
		    'Mailbox Tracking Log
		    LogText "Not processing Mail Submission Service Tracking Log File"
		    LogText now & vbtab & "SKIP" & vbtab & "MAIN 1 " & sLogFilePath & " Not processing Mail Submission Service Tracking Log File"		
	    End If
	  
	  End If
	  
  Else
    
    Err.Clear
    Set oFolderError = fso.GetFolder(sLogFilePath)
    If Err.number <> 0 Then
      LogText now & vbtab & "ERROR" & vbtab & "Error when attempting GetFolder(" & sLogFilePath & ") was '" & err.Description & "'"     
    End If
    Err.Clear
    Set oFileError = fso.GetFile(sLogFilePath)
    If Err.number <> 0 Then
      LogText now & vbtab & "ERROR" & vbtab & "Error when attempting GetFile(" & sLogFilePath & ") was '" & err.Description & "'"     
    End If
    LogText now & vbtab & "ERROR" & vbtab & "Path provided is not valid"
  
  End If
  
End If

LogText vbCRLF & now & vbtab & "INFO" & vbtab & "MAIN: Calling WriteSummary"
WriteSummary
CloseFiles
LogText vbCRLF & now & vbtab & "INFO" & vbtab & "MAIN: Calling ArchiveFiles"
ArchiveFiles
LogText vbCRLF & now & vbtab & "EXIT" & vbtab & "MAIN: Script done!"
'End Main Script


Function GetFolders(byRef sLogFilePath)
		On Error Resume Next
		LogText GetProcessInfo(FALSE)
		iLastFileSeqNum = 0
		sLastLogDateTimeEnd = NULL
		iFilesParsedOutOfSequenceCount = 0
		iEnterFunction = 0
		iTotalFolderFilesParsed = 0
		iMyNumFiles = iNumFiles
		GetFolders = FALSE
		Set oFolder = fso.GetFolder(sLogFilePath)
		If err.number <> 0 Then
  		LogText vbcrlf & now & vbtab & "ERROR" & vbtab & "GetFolders unable to access folder " & sLogFilePath & " (" & err.Description & ")"		  
  		LogText vbcrlf & now & vbtab & "ERROR" & vbtab & "GetFolders is exiting function"
  		Exit Function
		End If
		Set oFiles = oFolder.Files
		If err.number <> 0 Then
  		LogText vbcrlf & now & vbtab & "ERROR" & vbtab & "GetFolders unable to access files in " & sLogFilePath & " (" & err.Description & ")"		  
  		LogText vbcrlf & now & vbtab & "ERROR" & vbtab & "GetFolders is exiting function"
  		Exit Function
		End If
		iMyNumFiles=oFiles.count
		If iNumFiles = 0 Then 
			iMaxFiles = oFiles.count
		Else
			iMaxFiles = iNumFiles
		End If
		LogText vbcrlf & now & vbtab & "ENTER" & vbtab & "GetFolders working with Files in " & sLogFilePath & " (count=" & iMyNumFiles & ", maxParse=" & iMaxFiles & ")"
		iCurrentQueueCount = 0
		iCurrentQueueSizeBytes = 0
		iMaxQueueSizeBytes = 0
		iFileSeqNum = 0
		iPrevFileSeq = 0
		sPrevFilePrefix = ""		

		For each file in oFiles
			strCurrentFileName = file.name							
			If instr(1,lcase(strCurrentFileName),".log")>0 AND instr(1,lcase(strCurrentFileName),".idx")=0 Then 				
				sFilePrefix = left(strCurrentFileName,instr(1, strCurrentFileName,"-")-1)
				sFileSeqNum = mid(strCurrentFileName,instr(1, strCurrentFileName,"-")+1,instr(1, strCurrentFileName,".LOG")-instr(1, strCurrentFileName,"-")-1)
				iFileSeqNum = int(sFileSeqNum)
				For i = 1 to 5
					If len(sFileSeqNum) < 5 Then
						sFileSeqNum = "0" & sFileSeqNum
					End If
				Next
				'wscript.echo "adding " & sFilePrefix & "-" & sFileSeqNum
				oUnsortedFilesDictionary.Add sFilePrefix & "-" & sFileSeqNum,file.path
			End If
		Next
		
		aUnsortedFiles = oUnsortedFilesDictionary.Keys
		Set oSortedFiles = CreateObject("System.Collections.ArrayList" )
		For iElement = 0 To UBound(aUnsortedFiles)
			oSortedFiles.Add aUnsortedFiles(iElement)
		Next
		
		oSortedFiles.Sort	
		
		For iElement = 0 to oSortedFiles.Count - 1	
			strLogFilePath = oUnsortedFilesDictionary.Item(oSortedFiles(iElement))
			
			If instr(1,strLogFilePath,"MSGTRKM") = 0 Then
				'Edge/Hub Tracking Log
				iCurrentFileSeqNum = int(Right(oSortedFiles(iElement),5))
				sLastLogDateTimeEnd = sLogDateTimeEnd		
				ParseFile strLogFilePath				
				iLastFileSeqNum = iCurrentFileSeqNum
			Else
				'Mailbox Tracking Log
				'LogText now & vbtab & "SKIP" & vbtab & "GetFolders " & iEnterParseFileFunctionTotal & " " & strLogFilePath & " Not processing Mail Submission Service Tracking Log File"
				iCurrentFileSeqNum = int(Right(oSortedFiles(iElement),5))
				sLastLogDateTimeEnd = sLogDateTimeEnd		
				ParseFile strLogFilePath				
				iLastFileSeqNum = iCurrentFileSeqNum
			End If
			
			If iTotalFolderFilesParsed = iMaxFiles Then			
				Exit For
			Else
				'Wscript.echo "iNumFiles = " & iEnterFunction & " is less than " & iNumFiles
			End If			
		Next
		
		oUnsortedFilesDictionary.RemoveAll
				
		a = oServerLatencyDictionary.Keys
		iServerLatencyDictionaryCount = oServerLatencyDictionary.Count
		If iServerLatencyDictionaryCount > 0 Then
			For i = 0 to iServerLatencyDictionaryCount - 1
				WriteOrphanQueueEntry a(i), oServerLatencyDictionary.Item(a(i))
			Next
		End If

		LogText now & vbtab & "DONE" & vbtab & "GetFolders done working with Files in " & sLogFilePath & " (" & iTotalFolderFilesParsed & " files parsed, " & iTotalFilesParsed & " total)"
		LogText now & vbtab & "INFO" & vbtab & "GetFolders encountered " & iFilesParsedOutOfSequenceCount & " files parsed out of sequence (" & iTotalFilesParsedOutOfSequenceCount & " total parsed, " & iTotalFilesProcessOutOfSequenceCount & " total processed)"
		LogText vbCRLF & ">>> Removing " & iServerLatencyDictionaryCount & " oServerLatencyDictionary entries for " & sLocalServerName & vbCRLF
		oServerLatencyDictionary.RemoveAll
		
		Set oSubFolders = oFolder.SubFolders
		For each oSubFolder in oSubFolders
		
			Select Case ucase(strRole)
			
				Case "HUB"
					If instr(1,ucase(oSubFolder.path),"HUB")>0 Then
						GetFolders(oSubFolder.path)
					Else
						LogText now & vbtab & "SKIP" & vbtab & "GetFolders: Not " & strRole & " role, skipping Files in " & oSubFolder.path
					End If
					
				Case "EDGE"
					If instr(1,ucase(oSubFolder.path),"GWY")>0 Then
						GetFolders(oSubFolder.path)
					Else
						LogText now & vbtab & "SKIP" & vbtab & "GetFolders: Not " & strRole & " role, skipping Files in " & oSubFolder.path
					End If
					
				Case "DF"
					If instr(1,ucase(oSubFolder.path),"\DF-")>0 Then
						GetFolders(oSubFolder.path)
					Else
						LogText now & vbtab & "SKIP" & vbtab & "GetFolders: Not a dogfood server, skipping Files in " & oSubFolder.path
					End If
					
				Case "ALL"
					GetFolders(oSubFolder.path)
					
				Case Else
					LogText now & vbtab & "SKIP" & vbtab & "GetFolders: invalid role, skipping Files in " & oSubFolder.path

			End Select

		Next
				
		GetFolders = TRUE
	
End Function


Function ParseFile(sLogFilePath)
	On Error Resume Next
	Err.Clear
	iServerLatencyTotalCount = 0
	iDeliveryLatencyTotalCount = 0
	iServerLatencySLAMetCount = 0
	iDeliveryLatencySLAMetCount = 0
	iServerLatencyRecipientsNotCounted = 0
	iServerLatencyRecipientsCounted = 0
	iFileDuration = 0
	iFileReceiveStoreDriverCount = 0
	iFileDeliverCount = 0
	iFileReceiveSmtpCount = 0
	iFileSendSmtpCount = 0
	iFileFailCount = 0
	iFileDsnFailureCount = 0
	iFileDumpsterTrueCount = 0		
	iFileLastDeliveryCount = 0
	iFileTransferCount = 0
	iPercentDumpster = 0
	iFileDeferCount = 0
	iFileTotalMsgBytes = 0
	sRuntimeDiff = 0
	iEnterFunction = iEnterFunction + 1
	iEnterParseFileFunctionTotal = iEnterParseFileFunctionTotal + 1
	sLocalServerName = ""
	sLastLocalServerName = ""
	iSkipFileLine = 0
	iReadFileLine = 0
	iMaxFileQueueCount = iCurrentQueueCount
	iMaxQueueSizeBytes = iCurrentQueueSizeBytes
	sQueueExceptionStart = "01/01/1970 12:00:01 AM"
	iQueueExceptonDurationSec = 0
	iReSubmitEventCount = 0
	iResubmitDeliveryCount = 0
	bFileParsed = FALSE

	iLogFileEventCount = 0
	Set oFile = fso.GetFile(sLogFilepath)
	If err.number <> 0 Then
		LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " 0x" & hex(err.number) & ": " & replace(err.Description,vbcrlf," ")
		LogText now & vbtab & "EXIT" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Could not create file object"
		Exit Function
	End If
	If oFile.size=0 Then 
		LogText now & vbtab & "SKIP" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " is 0 bytes (" & iEnterFunction & " of " & iMyNumFiles & " files)"		
		Exit Function
	Else
		sStarted = now
		LogText now & vbtab & "OPEN" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " (" & iEnterFunction & " of " & iMyNumFiles & " files)"		
	End If
	Set oLogFile = fso.OpenTextFile(sLogFilepath,ForReading)
	If err.number <> 0 Then
		LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " 0x" & hex(err.number) & ": " & replace(err.Description,vbcrlf," ")
		LogText now & vbtab & "EXIT" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Could not open file"
		Exit Function		
	End If
    
	Do While oLogFile.AtEndOfStream <> True
							
		Err.Clear
		iParseFileDurationSec = datediff("s",sStarted,now)
		If iParseFileDurationSec > iParseFileDurationSecMax Then 
					LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Exceeded processing time of " & iParseFileDurationSecMax & " sec"				
					LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Last line processed was L" & iReadFileLine
					Exit Do
		End If
		sFileLine = oLogFile.Readline
		If err.number <> 0 Then
					LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " 0x" & hex(err.number) & ": " & replace(err.Description,vbcrlf," ")
					LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Last line parsed was L" & iReadFileLine
					Exit Do
		End If        
		iReadFileLine = iReadFileLine + 1
		if instr(1,left(sFileLine,1),"#") Then
										
			'#Date: 2007-05-20T22:39:46.784Z
			If instr(1,left(sFileLine,5),"#Date") Then
				sLogDateTime = GetDateTime(right(sFileLine,Len(sFileLine)-7))
				sLogDateTimeEnd = sLogDateTime
				bProcessFileOutofSequence = FALSE
				If iLastFileSeqNum = 0 Then
					LogText now & vbtab & "INFO" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " Processing the first file in folder (LogDate: " & sLogDateTime & " UTC)"
				ElseIf iLastFileSeqNum <> 0 AND iCurrentFileSeqNum <> iLastFileSeqNum + 1 Then							
					iPreviousLogFileDateDiff = DateDiff("d",sLastLogDateTimeEnd,sLogDateTime)							
					If iCurrentFileSeqNum=1 AND iPreviousLogFileDateDiff = 1 Then
						LogText now & vbtab & "INFO" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " Processing the first file in a new date sequence (LogDate: " & sLogDateTime & " UTC)"
					Else
						bProcessFileOutofSequence = TRUE
						If iPreviousLogFileDateDiff > 1 Then
							iTotalFilesProcessOutOfSequenceCount = iTotalFilesProcessOutOfSequenceCount + 1
							LogText now & vbtab & "WARNING" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " Processing has encountered a file out of sequence because date skipped (LogDate: " & sLogDateTime & " UTC, DateDiff=" & iPreviousLogFileDateDiff & ")"
						Else
							iMissingLogFiles = iCurrentFileSeqNum - iLastFileSeqNum - 1
							iTotalFilesProcessOutOfSequenceCount = iTotalFilesProcessOutOfSequenceCount + iMissingLogFiles
							LogText now & vbtab & "WARNING" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " Processing has encountered a file out of sequence for same date (" & iMissingLogFiles & " files missing)"
						End if
					End If							
				Else
					'Next file in sequence detected
				End If
				
				If DateDiff("d",sOldestLogDate,sLogDateTime) < 0 or (DateDiff("d",sOldestLogDate,sLogDateTime) = 0 and DateDiff("n",sOldestLogDate,sLogDateTime) < 0) Then 
					sOldestLogDate = sLogDateTime
					If bDebugOut Then LogText "OldestLogDate: " & sOldestLogDate
				End If
				If DateDiff("d",sNewestLogDate,sLogDateTime) > 0 or (DateDiff("d",sNewestLogDate,sLogDateTime) = 0 and DateDiff("n",sNewestLogDate,sLogDateTime) > 0)  Then 
					sNewestLogDate = sLogDateTime
					If bDebugOut Then LogText "sNewestLogDate: " & sNewestLogDate							
				End If
										
				If strFilterDate="" OR DateDiff("d",sLogDateTime,strFilterDate) = 0 Then
					iTotalFilesParsed = iTotalFilesParsed + 1
					iTotalFolderFilesParsed = iTotalFolderFilesParsed + 1
					bFileParsed = TRUE
					If bProcessFileOutofSequence Then
						iFilesParsedOutOfSequenceCount = iFilesParsedOutOfSequenceCount + 1
						iTotalFilesParsedOutOfSequenceCount = iTotalFilesParsedOutOfSequenceCount + 1								
					End If
				Else
					LogText now & vbtab & "DATE" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " Log file date (" & sLogDateTime & " UTC) is not equal " & strFilterDate & " (" & iEnterFunction & " of " & iMyNumFiles & " files)"		
					Exit Do
				End If
																						
			ElseIf instr(1,left(sFileLine,7),"#Fields") Then 
				sHeaderLine = sFileLine					
			Else
				'do nothing
			End If
		ElseIf instr(1,sFileLine,"#Software:") Then
			'do nothing
									
		Elseif instr(1,sFileLine,",") Then
        
					iEventCount = iEventCount + 1
					iLogFileEventCount = iLogFileEventCount + 1
	 									
					aFileLine = split(sFileLine,",")
					
					If ubound(aFileLine) <> 20 Then 
						
						iSplitError = iSplitError + 1
						
						'parsing file using ParseCsvLine takes 27 seconds compared with 2 seconds for split
						ParseCsvLine sFileLine, aFileLine
						If bDebugOut Then LogText "Return from ParseCsvLine with array size: " & ubound(aRegExFileLine)

						If ubound(aFileLine) <> 20 Then
							iSkipFileLine = iSkipFileLine + 1
							LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Skipping invalid line on L" & iReadFileLine & " (" & iSkipFileLine & ")"
							If iSkipFileLine = iSkipFileLineMax Then 
								LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Exceeded max invalid line count of " & iSkipFileLineMax
								LogText now & vbtab & "ERROR" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & sLogFilepath & " Closing File"
								Exit Do
							End If
						Else
							ExtractEventInfo aFileLine
							sLogDateTimeEnd = GetDateTime(aFileLine(iEventTime))
					    
							If err.number <> 0 Then
								iErrorEvent = iErrorEvent + 1			
								if bLogError Then oLogError.writeline sFileLine		
								err.Clear				
							End If
					    
		
						End If						
						
												
					Else
						
						ExtractEventInfo aFileLine											
						sLogDateTimeEnd = GetDateTime(aFileLine(iEventTime))
					
						If err.number <> 0 Then
							iErrorEvent = iErrorEvent + 1			
							if bLogError Then oLogError.writeline sFileLine		
							err.Clear				
						End If
				
		        					
					End If
            
		Else
					'Wscript.echo "*** Exit Parse Loop for " & sLogFilePath
					Exit Function
					
		End If
	Loop
	oLogFile.close
	sParseEnd = now
	sRuntimeDiff = datediff("s",sStarted,sParseEnd)
		LogText now & vbtab & "CLOSE" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & iLogFileEventCount & " Events Processed in " & sRuntimeDiff & " sec"
    
	If bFileParsed Then
		iFileDuration = datediff("s",sLogDateTime,sLogDateTimeEnd)/60 '+ datediff("n",sLogDateTime,sLogDateTimeEnd)*60)/60
		If iFileDuration > iFileDurationMax Then iFileDurationMax = round(iFileDuration,1)
		If iFileDuration < iFileDurationMin Then iFileDurationMin = round(iFileDuration,1)
		iTotalFileDateTimeDiffMin = iTotalFileDateTimeDiffMin + iFileDuration
    
		a = oServerLatencyDictionary.Keys
		iServerLatencyDictionaryEntriesRemoved = 0
		iServerLatencyDictionaryCount = oServerLatencyDictionary.Count
		If iServerLatencyDictionaryCount > 0 Then
			For i = 0 to iServerLatencyDictionaryCount - 1
				strUniqueMessageId = a(i)
				aLatencyTrackingEntry = oServerLatencyDictionary.Item(a(i))
				If bPruneDictionary Then
					If aLatencyTrackingEntry(6) = "<>" Then
						iServerLatencyDictionaryEntriesRemoved = iServerLatencyDictionaryEntriesRemoved + 1
						iServerLatencyRecipientsNotCounted = iServerLatencyRecipientsNotCounted + int(aLatencyTrackingEntry(1))						
						TrackLastDelivery aLatencyTrackingEntry, "CLEANUP", ""
						'Wscript.echo "Removing " & i & " " & a(i) & " (null return)"
						oServerLatencyDictionary.Remove a(i)
					ElseIf aLatencyTrackingEntry(3) = "ROUTING:TRANSFER(Resolver)" Then
						'Resolver bifurcation for DG's that suppress DSN's
						iServerLatencyDictionaryEntriesRemoved = iServerLatencyDictionaryEntriesRemoved + 1
						iServerLatencyRecipientsNotCounted = iServerLatencyRecipientsNotCounted + int(aLatencyTrackingEntry(1))						
						TrackLastDelivery aLatencyTrackingEntry, "CLEANUP", ""
						'Wscript.echo "Removing " & i & " " & a(i) & " (null return)"
						oServerLatencyDictionary.Remove a(i)					
					ElseIf (aLatencyTrackingEntry(2)>0 AND aLatencyTrackingEntry(7)) Then
						'iServerLatencyDictionaryEntriesRemoved = iServerLatencyDictionaryEntriesRemoved + 1
						'iServerLatencyRecipientsNotCounted = iServerLatencyRecipientsNotCounted + int(aLatencyTrackingEntry(1))
						'Wscript.echo "Removing " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)
						'TrackLastDelivery aLatencyTrackingEntry, "CLEANUP"
						'Wscript.echo "Attempting to remove " & a(i) & " (delivery or transfer > 0)"
						'oServerLatencyDictionary.Remove a(i)
					Else
						'If i < 10 Then Wscript.echo "Valid oServerLatencyDictionary entry: " & a(i) & vbtab & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5) & vbtab & aLatencyTrackingEntry(6)						
					End If		
				End If		
			Next
		End If
		
		'Compute Rates for events per File processed
		iAvgFailRatePerMin = round(iFileFailCount/iFileDuration,1)
		iAvgDeferRatePerMin = round(iFileDeferCount/iFileDuration,1)
		iAvgTransferRatePerMin = round(iFileTransferCount/iFileDuration,1)
		iAvgDsnFailRatePerMin = round(iFileDsnFailureCount/iFileDuration,1)	
		If iFileReceiveSmtpCount > 0 OR iFileReceiveStoreDriverCount > 0 Then
			iFileAvgMsgBytes = round(iFileTotalMsgBytes/(iFileReceiveSmtpCount + iFileReceiveStoreDriverCount),0)	
		Else
			iFileAvgMsgBytes = 0
		End If
		If iFileLastDeliveryCount > 0 Then
			iPercentDumpster = round(100*iFileDumpsterTrueCount/iFileLastDeliveryCount,1)	
		Else
			iPercentDumpster = 0
		End If
		iAvgStoreDriverReceiveRatePerMin = round(iFileReceiveStoreDriverCount/iFileDuration,1)
		iAvgSmtpReceiveRatePerMin = round(iFileReceiveSmtpCount/iFileDuration,1)
		iAvgSmtpSendRatePerMin = round(iFileSendSmtpCount/iFileDuration,1)
		iAvgStoreDriverDeliverRatePerMin = round(iFileDeliverCount/iFileDuration,1)
		If iAvgFailRatePerMin > iMaxAvgFailPerMin Then iMaxAvgFailPerMin = iAvgFailRatePerMin
		If iAvgDsnFailRatePerMin > iMaxAvgFileDsnCount Then iMaxAvgFileDsnCount = iAvgDsnFailRatePerMin
				
		iPercentServerLatencyRecipientsCounted = round(100*iServerLatencyRecipientsCounted/(iServerLatencyRecipientsCounted+iServerLatencyRecipientsNotCounted),2)
		iPercentServerLatencySLAMet = round(100*iServerLatencySLAMetCount/iServerLatencyTotalCount,2)
		If iDeliveryLatencyTotalCount > 0 Then
			iPercentDeliveryLatencySLAMet = round(100*iDeliveryLatencySLAMetCount/iDeliveryLatencyTotalCount,2)
		Else
			iPercentDeliveryLatencySLAMet = 0
		End If
		iAllServerLatencyTotalCount = iAllServerLatencyTotalCount + iServerLatencyTotalCount

		iMsgIdCount = iMsgIdCount + oMessageIdDictionary.Count
		If oMessageIdDictionary.Count > 0 Then iAvgLogEventPerMsgId = iLogFileEventCount/oMessageIdDictionary.Count
		If iAvgLogEventPerMsgId > iMaxLogEventPerMsgId Then iMaxLogEventPerMsgId = iAvgLogEventPerMsgId
		If iAvgLogEventPerMsgId < iMinLogEventPerMsgId Then iMinLogEventPerMsgId = iAvgLogEventPerMsgId
		If oMessageIdDictionary.Count > 0 Then 
			oLogStatistics.writeline sLogFilePath & "," & ucase(sLocalServerName) & "," & iLogFileEventCount & "," & (iLogFileEventCount/oMessageIdDictionary.Count) & "," & sLogDateTime & "," & sLogDateTimeEnd & "," & round(iFileDuration,1) & "," & iAvgStoreDriverReceiveRatePerMin & "," & iAvgSmtpReceiveRatePerMin & "," & iAvgTransferRatePerMin & "," & iAvgStoreDriverDeliverRatePerMin & "," & iAvgSmtpSendRatePerMin & "," & iAvgDsnFailRatePerMin & "," & iAvgFailRatePerMin & "," & iAvgDeferRatePerMin & "," & iServerLatencyTotalCount & "," & iPercentServerLatencySLAMet & "," & iPercentServerLatencyRecipientsCounted & "," & iDeliveryLatencyTotalCount & "," & iPercentDeliveryLatencySLAMet & "," & iPercentDumpster & "," & iReSubmitEventCount & "," & iResubmitDeliveryCount & "," & iFileAvgMsgBytes & "," & iMaxFileQueueCount & "," & iQueueExceptonDurationSec & "," & iMaxQueueSizeBytes & "," & sStarted & "," & sRuntimeDiff 
			'Wscript.echo "Queue Size = " & iCurrentQueueCount & vbtab & "Max File = " & iMaxFileQueueCount
		End If
		oMessageIdDictionary.RemoveAll
		iParseFileLatencyAfterClose = datediff("s",sParseEnd,now)
		LogText now & vbtab & "EXIT" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal & " " & iServerLatencyDictionaryEntriesRemoved & " entries removed from ServerLatencyDictionary with size " & iServerLatencyDictionaryCount & " (" & oServerLatencyDictionary.Count & ") in " & iParseFileLatencyAfterClose & " sec"
	Else
		LogText now & vbtab & "EXIT" & vbtab & "ParseFile " & iEnterParseFileFunctionTotal
	End If
End Function

Function ExtractEventInfo(byRef aFileLine)
	On Error Resume Next
	bServerLatencySLAMet = FALSE
	bDeliveryLatencyMet = FALSE
	
	If aFileLine(iMessageId) = "" Then
		
		Select Case aFileLine(iSource)
					
			Case "ADSN" 
				'Wscript.echo "No Message Id: " & sFileLine	
				If oEmptyMessageidDictionary.Exists(aFileLine(iInternalMsgId)) Then
					'Wscript.echo "No Message Id: InternalMessageId " & aFileLine(iInternalMsgId) & " already exists in dictionary for ADSN"
				Else
					oEmptyMessageidDictionary.Add aFileLine(iInternalMsgId), int(aFileLine(iRecipCount))
				End If
									
			Case "MAILBOXRULE" 
				'Wscript.echo "No Message Id: " & sFileLine	
				If oEmptyMessageidDictionary.Exists(aFileLine(iInternalMsgId)) Then
					'Wscript.echo "No Message Id: InternalMessageId " & aFileLine(iInternalMsgId) & " already exists in dictionary for MAILBOXRULE"
				Else
					oEmptyMessageidDictionary.Add aFileLine(iInternalMsgId), int(aFileLine(iRecipCount))
				End If
					
			Case Else
				
				If oEmptyMessageidDictionary.Exists(aFileLine(iInternalMsgId)) Then
					
					iBlankMsgRecipCount = oEmptyMessageidDictionary.Item(aFileLine(iInternalMsgId)) 					
					If iBlankMsgRecipCount <= int(aFileLine(iRecipCount)) Then
						oEmptyMessageidDictionary.Remove aFileLine(iInternalMsgId)
					Else
						oEmptyMessageidDictionary.Item(aFileLine(iInternalMsgId)) = iBlankMsgRecipCount - int(aFileLine(iRecipCount))
					End If
										
				ElseIf aFileLine(iEventId) = "TRANSFER" Then
					
					If oEmptyMessageidDictionary.Exists(aFileLine(iReference)) Then						
					
						'Remove recipients from reference
						iBlankMsgRecipCount = oEmptyMessageidDictionary.Item(aFileLine(iReference)) 
						If iBlankMsgRecipCount <= int(aFileLine(iRecipCount)) Then
							oEmptyMessageidDictionary.Remove aFileLine(iReference)
						Else
							oEmptyMessageidDictionary.Item(aFileLine(iReference)) = iBlankMsgRecipCount - int(aFileLine(iRecipCount))
						End If
						
						'Add new message created through transfer to dictionary
						If oEmptyMessageidDictionary.Exists(aFileLine(iInternalMsgId)) Then
							Wscript.echo "No Message Id: InternalMessageId " & aFileLine(iInternalMsgId) & " already exists in dictionary for ADSN"
						Else
							oEmptyMessageidDictionary.Add aFileLine(iInternalMsgId), int(aFileLine(iRecipCount))
						End If					
					
					Else
						'Wscript.echo "No Message Id (" & aFileLine(iSource) & ":" & aFileLine(iEventId) & ") - InternalMessageId " & aFileLine(iReference) & " could not be found in dictionary for InternalMsgId " & aFileLine(iInternalMsgId)
					End If
				
				Else
				
					'Wscript.echo "No Message Id (" & aFileLine(iSource) & ":" & aFileLine(iEventId) & ") - InternalMessageId " & aFileLine(iInternalMsgId) & " could not be found in dictionary"
					
				End If
		
		End Select
		
	End If
	
	If oMessageIdDictionary.Exists(aFileLine(iMessageId)) Then
		iMsgIdEventCount = oMessageIdDictionary.Item(aFileLine(iMessageId))
		oMessageIdDictionary.Item(aFileLine(iMessageId)) = iMsgIdEventCount + 1
	Else
		oMessageIdDictionary.Add aFileLine(iMessageId), 1
	End if
	
	If aFileLine(iEventID) <> "SUBMIT" Then
		If UpdateServerLatency(aFileLine, bServerLatencySLAMet) Then
		Else
			LogText now & vbtab & "ERROR" & vbtab & "UpdateServerLatency returned FALSE for " & sFileLine		
		End If
	End If
	
	sCurrentEventTime=GetDateTime(aFileLine(iEventTime))
	'Default (overriden in the events that use iClientName for local server)
	sLastLocalServerName = sLocalServerName
	sLocalServerName = aFileLine(iServerName)
	  
	Select Case aFileLine(iEventID)
		
		Case "RECEIVE"
			
			iReceiveCount = iReceiveCount + 1
			iTotalMsgBytes = iTotalMsgBytes + int(aFileLine(iTotalBytes))
			iFileTotalMsgBytes = iFileTotalMsgBytes + int(aFileLine(iTotalBytes))
						
			If aFileLine(iSource) = "SMTP" Then
				sReceiveClientDictionaryKey = aFileLine(iServerName) & ":" & replace("SMTP(" & aFileLine(iConnectorId),":","_") & "):" & aFileLine(iClientIP)
			ElseIf aFileLine(iSource) = "AGENT" Then
				 sReceiveClientDictionaryKey = aFileLine(iServerName) & ":" & aFileLine(iSource) & "(" & replace(aFileLine(iSourceContext),":","_") & "):" & aFileLine(iClientName)
				 'Wscript.echo sReceiveClientDictionaryKey & vbtab & aFileLine(iConnectorId) 
			Else
				sReceiveClientDictionaryKey = aFileLine(iServerName) & ":" & aFileLine(iSource) & ":" & aFileLine(iClientName)
			End If				
				
			If oReceiveClientDictionary.Exists(sReceiveClientDictionaryKey) Then
				aExistReceiveClientDictionaryItem = oReceiveClientDictionary.Item(sReceiveClientDictionaryKey)
				aExistReceiveClientDictionaryItem(0) = aExistReceiveClientDictionaryItem(0) + 1
				aExistReceiveClientDictionaryItem(1) = aExistReceiveClientDictionaryItem(1) + int(aFileLine(iTotalBytes))				
				oReceiveClientDictionary.Item(sReceiveClientDictionaryKey) = aExistReceiveClientDictionaryItem
			Else
				Redim aAddReceiveClientDictionaryItem(1)
				aAddReceiveClientDictionaryItem(0) = 1
				aAddReceiveClientDictionaryItem(1) = int(aFileLine(iTotalBytes))
				oReceiveClientDictionary.Add sReceiveClientDictionaryKey,aAddReceiveClientDictionaryItem
			End If
																						
			Select Case aFileLine(iSource)
				Case "SMTP"
						iReceiveSmtpRecipCount = iReceiveSmtpRecipCount + int(aFileLine(iRecipCount))
						iReceiveSmtpCount = iReceiveSmtpCount + 1
						iFileReceiveSmtpCount = iFileReceiveSmtpCount + 1
						
				Case "STOREDRIVER"
						iReceiveSDRecipCount = iReceiveSDRecipCount + int(aFileLine(iRecipCount))
						iReceiveStoreDriverCount = iReceiveStoreDriverCount + 1
						iFileReceiveStoreDriverCount = iFileReceiveStoreDriverCount + 1
						
						strSenderDomain = lcase(right(aFileline(iSender),len(aFileLine(iSender))-instr(1,aFileLIne(iSender),"@")))
						if oSDSubmitDomainDictionary.Exists(strSenderDomain) Then															
							iCount = oSDSubmitDomainDictionary.Item(strSenderDomain)
							oSDSubmitDomainDictionary.Item(strSenderDomain) = iCount + 1
						Else
							oSDSubmitDomainDictionary.Add strSenderDomain, 1
						End If
						
						'Create dictionary to count top senders (internal only)
						sTopSendersbySubmitDictionaryKey = aFileLine(iClientName) & ":" & aFileLine(iSender) 
						If oTopSendersbySubmitDictionary.Exists(sTopSendersbySubmitDictionaryKey) Then
							iTopSenderCount = int(oTopSendersbySubmitDictionary.Item(sTopSendersbySubmitDictionaryKey))
							oTopSendersbySubmitDictionary.Item(sTopSendersbySubmitDictionaryKey) = iTopSenderCount + 1
						Else
							oTopSendersbySubmitDictionary.Add sTopSendersbySubmitDictionaryKey, 1
						End If
						
				Case Else
					
			End Select
			
			If instr(1,aFileLIne(iSender),"PUBStore") or instr(1,aFileLIne(iSender),"-IS@") Then 
				If oPFSubjectDictionary.Exists(aFileLine(iSubject)) Then
					iPFCount = oPFSubjectDictionary.Item(aFileLine(iSubject))
					oPFSubjectDictionary.Item(aFileLine(iSubject)) = iPFCount + 1
				Else
					oPFSubjectDictionary.add aFileLine(iSubject),1
				End If
			End If
			
			If (0 <= int(aFileLine(iTotalBytes))) and (int(aFileLine(iTotalBytes)) <= 1024) Then
					aMsgSize(0,1) = int(aMsgSize(0,1)) + 1
			Else
				For i = 1 to 16 
					If (2^(i-1)*1024 < int(aFileLine(iTotalBytes))) and (int(aFileLine(iTotalBytes)) <= (2^i)*1024) Then
						aMsgSize(i,1) = int(aMsgSize(i,1)) + 1
					End If
				Next				                
			End If
			If int(aFileLine(iTotalBytes)) > iMaxMessageSizeThresholdKB*1024 Then 
				If iExceedsMaxMessageSize = 0 Then
					oMessageSizeExceptionResults.writeline sHeaderLine
				End If
				iExceedsMaxMessageSize = iExceedsMaxMessageSize + 1
				oMessageSizeExceptionResults.writeline sFileLine
			End If
								
			sTimeReceived = aFileLine(iEventTime)
      
		Case "SEND"
			sLocalServerName = aFileLine(iClientName)
			iSendSmtpRecipCount = iSendSmtpRecipCount + int(aFileLine(iRecipCount))
			iSendCount = iSendCount + 1
			iFileSendSmtpCount = iFileSendSmtpCount + 1
				
			If TrackNextHopLatency(aFileLine, bServerLatencySLAMet, bDeliveryLatencyMet) Then
			Else
				LogText now & vbtab & "ERROR" & vbtab & "TrackNextHopLatency failed, exit Function ExtractEventInfo"
				Exit Function
			End If
		
		Case "EXPAND"
			iExpandCount = iExpandCount + 1
			iExpandRecipCount = iExpandRecipCount + int(aFileLine(iRecipCount))		
			if iMaxExpandRecipCount < int(aFileLine(iRecipCount)) then iMaxExpandRecipCount = int(aFileLine(iRecipCount))

			sCurrentEventTime=GetDateTime(aFileLine(iEventTime))
			strUniqueMessageid = aFileLine(iInternalMsgId) & "," & aFileLine(iMessageId)
			If oServerLatencyDictionary.Exists(strUniqueMessageId) Then
				aLatencyTrackingEntry = oServerLatencyDictionary.Item(strUniqueMessageid)
				iExpandLatencySec = datediff("s",aLatencyTrackingEntry(0),sCurrentEventTime)
			Else
				iExpandLatencySec = 0
			End If								
			
			If oExpandedDLDictionary.Exists(aFileLine(iRelatedRecipAddr)) Then				
				aExpandDetail = oExpandedDLDictionary.Item(aFileLine(iRelatedRecipAddr))
				aExpandDetail(1) = aExpandDetail(1) + 1
				aExpandDetail(2) = aExpandDetail(2) + iExpandLatencySec
				'Wscript.echo "Updated DG " & aFileLine(iRelatedRecipAddr) & " with RecipCount " & aExpandDetail(0)	& " and ExpandCount " & aExpandDetail(1)
				oExpandedDLDictionary.Item(aFileLine(iRelatedRecipAddr)) = aExpandDetail
			Else
				ReDim aExpandDetail(2) '0=RecipCount,1=ExpandCount,2=TotalExpandLatency
				iUniqueDL = iUniqueDL + 1
				aExpandDetail(0) = int(aFileLine(iRecipCount))	
				aExpandDetail(1) = 1
				aExpandDetail(2) = iExpandLatencySec
				'Track latency
				'Wscript.echo "Adding DG " & aFileLine(iRelatedRecipAddr) & " with RecipCount " & aExpandDetail(0) & " and Expandcount " & aExpandDetail(1)
				oExpandedDLDictionary.Add aFileLine(iRelatedRecipAddr), aExpandDetail
			End If
												
		Case "DELIVER"
			sLocalServerName = aFileLine(iClientName)
			'Count events
			iSdDeliverRecipCount = iSdDeliverRecipCount + int(aFileLine(iRecipCount))
			iDeliverCount = iDeliverCount + 1
			iFileDeliverCount = iFileDeliverCount + 1			
			
			If aFileLine(iReturnPath)="<>" Then
				If instr(1,lcase(aFileLine(iSender)),sInternalDNSSuffix)>0 AND instr(1,lcase(aFileLine(iSender)),"microsoftexchange329e71ec88ae4615bbc36ab6ce41109e")>0 Then
					iMERDeliverCount = iMERDeliverCount + 1
				Else
					If instr(1,aFileLine(iMessageId),sInternalDNSSuffix) Then
						'Internal MessageId found
						iInternalNullReversePathDeliverCount = iInternalNullReversePathDeliverCount  +1
					Else
						'External MessageId found
						iExternalNullReversePathDeliverCount = iExternalNullReversePathDeliverCount  +1
					End IF
				End If
			End If

			If aFileLine(iRecipCount) = 1 Then
				aRecipients = Array(aFileLine(iRecipients))
			Else
				aRecipients = split(aFileLine(iRecipients),";")
			End If
			
			'If ubound(aRecipients)>0 Then Wscript.echo "Found " & ubound(aRecipients) & " Recipients (RecipCount=" & aFileline(iRecipCount) & ")"
			For i = 0 to ubound(aRecipients)
				'Create dictionary to count top recipients (internal only)
				sTopRecipientsDictionaryKey = aFileLine(iServerName) & ":" & aRecipients(i)
				If oTopRecipientsDictionary.Exists(sTopRecipientsDictionaryKey) Then
					iTopSenderCount = int(oTopRecipientsDictionary.Item(sTopRecipientsDictionaryKey))
					oTopRecipientsDictionary.Item(sTopRecipientsDictionaryKey) = iTopSenderCount + 1
					'If ubound(aRecipients)>0 Then Wscript.echo "DELIVER to " & aRecipients(i) & " (count = " & iTopSenderCount + 1 & ")"
				Else
					'If ubound(aRecipients)>0 Then Wscript.echo "DELIVER to " & aRecipients(i)
					oTopRecipientsDictionary.Add sTopRecipientsDictionaryKey, 1
				End If
			Next
			
			'Create dictionary to count top senders (based on DELIVER events, includes both internal and external senders)
			If oTopSendersbyDeliverDictionary.Exists(aFileLine(iSender)) Then
				iTopSenderCount = int(oTopSendersbyDeliverDictionary.Item(aFileLine(iSender)))
				oTopSendersbyDeliverDictionary.Item(aFileLine(iSender)) = iTopSenderCount + 1
			Else
				oTopSendersbyDeliverDictionary.Add aFileLine(iSender), 1
			End If
			
			'Parse MessageInfo field to calculate end-to-end delivery
			strOriginalArrivalTime = GetOriginalArrivalTime(aFileLine(iMessageInfo))				
			strDeliverTime = GetDateTime(aFileLine(iEventTime))
			If strOriginalArrivalTime<>"" AND strDeliverTime<>"" Then
				iDeliveryLatency = DateDiff("s",strOriginalArrivalTime,strDeliverTime)
			End If					
			If aFileLine(iMessageInfo)<>"" and iDeliveryLatency=>0 Then
				iDeliveryLatencyTotalCount = iDeliveryLatencyTotalCount + 1
				If iDeliveryLatency > iDeliveryLatencySLA Then
					If iDeliveryLatency > iDeliveryLatencySLAExceptionLoggingThreshold Then
						sDeliveryLatencyExceptionDictionaryKey = aFileLine(iMessageId)
						If oDeliveryLatencyExceptionDictionary.Exists(sDeliveryLatencyExceptionDictionaryKey) Then
							aExistDeliveryLatencyExceptionDictionaryItem = oDeliveryLatencyExceptionDictionary.Item(sDeliveryLatencyExceptionDictionaryKey)
							aExistDeliveryLatencyExceptionDictionaryItem(3) = aExistDeliveryLatencyExceptionDictionaryItem(3) + int(aFileLine(iRecipCount))
							aExistDeliveryLatencyExceptionDictionaryItem(4) = aExistDeliveryLatencyExceptionDictionaryItem(4) + 1
							If aExistDeliveryLatencyExceptionDictionaryItem(5) > iDeliveryLatency Then aExistDeliveryLatencyExceptionDictionaryItem(5) = iDeliveryLatency
							If aExistDeliveryLatencyExceptionDictionaryItem(6) < iDeliveryLatency Then aExistDeliveryLatencyExceptionDictionaryItem(6) = iDeliveryLatency
							If NOT(instr(1,aExistDeliveryLatencyExceptionDictionaryItem(7),aFileLine(iClientName))>0) Then
								aExistDeliveryLatencyExceptionDictionaryItem(7) = aExistDeliveryLatencyExceptionDictionaryItem(7) & ";" & aFileLine(iClientName)
							End If								
							oDeliveryLatencyExceptionDictionary.Item(sDeliveryLatencyExceptionDictionaryKey) = aExistDeliveryLatencyExceptionDictionaryItem
						Else
							ReDim aAddDeliveryLatencyExceptionDictionaryItem(7) 'strOriginalArrivalTime,iClientName,iTotalBytes,iSender,iRecipCount,iDeliverCount,iMinDeliveryLatency,iMaxDeliveryLatency
							aAddDeliveryLatencyExceptionDictionaryItem(0) = strOriginalArrivalTime
							aAddDeliveryLatencyExceptionDictionaryItem(1) = aFileLine(iTotalBytes)
							aAddDeliveryLatencyExceptionDictionaryItem(2) = aFileLine(iSender)
							aAddDeliveryLatencyExceptionDictionaryItem(3) = aFileLine(iRecipCount)
							aAddDeliveryLatencyExceptionDictionaryItem(4) = 1
							aAddDeliveryLatencyExceptionDictionaryItem(5) = iDeliveryLatency
							aAddDeliveryLatencyExceptionDictionaryItem(6) = iDeliveryLatency
							aAddDeliveryLatencyExceptionDictionaryItem(7) = aFileLine(iClientName)
							oDeliveryLatencyExceptionDictionary.Add sDeliveryLatencyExceptionDictionaryKey,aAddDeliveryLatencyExceptionDictionaryItem
						End If
					End If
				Else
					bDeliveryLatencyMet = TRUE
					iDeliveryLatencySLAMetCount = iDeliveryLatencySLAMetCount + 1							
				End If
				If 0 =< iDeliveryLatency and iDeliveryLatency <= 1 Then
						aMsgDeliveryLatency(0,1) = int(aMsgDeliveryLatency(0,1)) + 1
				Else
					For i = 1 to 24 
						If 2^(i-1) < iDeliveryLatency and iDeliveryLatency <= (2^i) Then
							aMsgDeliveryLatency(i,1) = int(aMsgDeliveryLatency(i,1)) + 1
						End If
					Next				                
				End If
									
			ElseIf aFileLine(iMessageInfo)<>"" and iDeliveryLatency<0 Then
				bDeliveryLatencyMet = TRUE
				iDeliveryLatencyCountLtZero = iDeliveryLatencyCountLtZero + 1
			
			ElseIf aFileLine(iMessageInfo)="" Then
				'Empty OriginalArrivalTime
				bDeliveryLatencyMet = TRUE
				iDeliveryLatencyCountUnknown = iDeliveryLatencyCountUnknown + 1
				
			Else
				bDebugOut=TRUE
				ParseCsvLine sFileLine, aFileLine
				LogText now & vbtab & "ERROR" & vbtab & "Return from ParseCsvLine with array size: " & ubound(aFileLine)	
				LogText now & vbtab & "ERROR" & vbtab & "t1:" & aFileLine(iMessageInfo) & vbtab & strOriginalArrivalTime 
				LogText now & vbtab & "ERROR" & vbtab & "t2:" & aFileLine(iEventTime) & vbtab & strDeliverTime 						
				LogText now & vbtab & "ERROR" & vbtab & "event: " & sFileLine
			End If
			
			If TrackNextHopLatency(aFileLine, bServerLatencySLAMet, bDeliveryLatencyMet) Then
			Else
				LogText now & vbtab & "ERROR" & vbtab & "TrackNextHopLatency failed"
				LogText now & vbtab & "ERROR" & vbtab & "event: " & sFileLine
			End If

        
		Case "FAIL"
			sLocalServerName = aFileLine(iClientName)
			iFileFailCount = iFileFailCount + 1
			iFailCount = iFailCount + 1
			iFailRecipientCount = iFailRecipientCount + aFileLine(iRecipCount)
			sFailReason = aFileLine(iRecipientStatus)
			'Wscript.echo "call FixResponseCode" 
			sFailReason=FixResponseCode(sFailReason)
							
			If int(aFileLine(iRecipCount)) > 1 Then 
				aFailReason=split(sFailReason,";")				
				strFailReason = aFailReason(0)
			Else
				strFailReason=sFailReason
			End If
			
			If sFailReason<>"" Then 
				
				If instr(1,sFailReason,"STOREDRV") Then
					If instr(1,strFailReason,". The following information should help identify the cause") > 0 Then 
						strFailReason = left(strFailReason,instr(1,strFailReason,". The following information should help identify the cause")-1)
						
						If instr(1,sFailReason,"MapiExceptionShutoffQuotaExceeded") Then
							aRecipients = split(aFileLine(iRecipients),";")
							For i = 0 to ubound(aRecipients) 
								If oMailboxFullDictionary.Exists(aRecipients(i)) Then		
									iFullCount = oMailboxFullDictionary.Item(aRecipients(i))
									oMailboxFullDictionary.Item(aRecipients(i)) = iFullCount + 1
								Else
									oMailboxFullDictionary.Add aRecipients(i),1
								End If
							Next
						End If							
					End If 
					'End STOREDRV
					
				ElseIf instr(1,sFailReason,"RESOLVER.ADR") Then
					If instr(1,sFailReason,"ExRecipNotFound") Then
						aRecipients = split(aFileLine(iRecipients),";")
						For i = 0 to ubound(aRecipients) 
							If instr(1,aRecipients(i),"IMCEAEX-") Then
								aRecipients(i) = replace(aRecipients(i),"_ou=","_OU=")
								aRecipients(i) = replace(aRecipients(i),"_o=","_O=")
								aRecipients(i) = replace(aRecipients(i),"_cn=","_CN=")
								'Wscript.echo "aRecipients(" & i & ") = " & aRecipients(i)
								strImcEaValue = mid(aRecipients(i),1,instr(5,aRecipients(i),"-")-1) & ":" & mid(aRecipients(i),instr(1,aRecipients(i),"_O=")+3,instr(1,aRecipients(i),"_OU=")-instr(1,aRecipients(i),"_O=")-3)
							ElseIf instr(1,aRecipients(i),"IMCEA") Then
								strImcEaValue = mid(aRecipients(i),1,instr(5,aRecipients(i),"-")-1)
							Else
								'Wscript.echo aRecipients(i) & " not found"
							End If
							If oIMCEADictionary.Exists(strImcEaValue) Then		
								iImcEaCount = oIMCEADictionary.Item(strImcEaValue)
								oIMCEADictionary.Item(strImcEaValue) = iImcEaCount + 1
							Else
								oIMCEADictionary.Add strImcEaValue,1
							End If
						Next
					ElseIf instr(1,sFailReason,"RecipNotFound") Then			
						aRecipients = split(aFileLine(iRecipients),";")
						For i = 0 to ubound(aRecipients) 
							iRecipNotFoundCountTotal = iRecipNotFoundCountTotal + 1
							If oRecipNotFoundDictionary.Exists(aRecipients(i)) Then		
								iRecipNotFoundCount = oRecipNotFoundDictionary.Item(aRecipients(i))
								oRecipNotFoundDictionary.Item(aRecipients(i)) = iRecipNotFoundCount + 1
							Else
								oRecipNotFoundDictionary.Add aRecipients(i),1
							End If
						Next
					End If
					'End RESOLVER.ADR																									
					
				ElseIf instr(1,sFailReason,"RESOLVER") Then
					'Wscript.echo sFailReason
				
				ElseIf instr(1,sFailReason,"QUEUE.Expired") Then
					If instr(1,aFileLine(iRecipients),";") Then
						aRecipients = split(aFileLine(iRecipients),";")
						strRecipient = aRecipients(0)
					Else
						strRecipient = aFileLine(iRecipients)
					End If
					
					sCurrentEventTime = GetDateTime(aFileLine(iEventTime)) 

					strDomainExpired = lcase(right(strRecipient,len(strRecipient)-instr(1,strRecipient,"@")))
					'Wscript.echo strDomainExpired
										
					If oDomainExpiredDictionary.Exists(strDomainExpired) Then						
						aDomainExpiredDictionaryEntry = oDomainExpiredDictionary.Item(strDomainExpired)
						aDomainExpiredDictionaryEntry(0) = aDomainExpiredDictionaryEntry(0) + int(aFileLine(iRecipCount))
						aDomainExpiredDictionaryEntry(1) = aDomainExpiredDictionaryEntry(1) + 1
						If aDomainExpiredDictionaryEntry(4) = sBlockDomainDate Then
							aDomainExpiredDictionaryEntry(3) = 0
						Else
							aDomainExpiredDictionaryEntry(3) = aDomainExpiredDictionaryEntry(3) + datediff("s",aDomainExpiredDictionaryEntry(4),sCurrentEventTime)
						End If
						aDomainExpiredDictionaryEntry(4) = sCurrentEventTime
						oDomainExpiredDictionary.Item(strDomainExpired) = aDomainExpiredDictionaryEntry
					Else
						Redim aDomainExpiredDictionaryEntry(4) '0=RecipCount,1=MsgCount,2=OrigDateTime,3=TotalSec,4=LastDateTime
						aDomainExpiredDictionaryEntry(0) = int(aFileLine(iRecipCount)) 
						aDomainExpiredDictionaryEntry(1) = 1						
						aDomainExpiredDictionaryEntry(2) = sCurrentEventTime
						aDomainExpiredDictionaryEntry(3) = 0
						aDomainExpiredDictionaryEntry(4) = sCurrentEventTime
						oDomainExpiredDictionary.Add strDomainExpired, aDomainExpiredDictionaryEntry
						
					End If
				
				ElseIf instr(1,sFailReason,"ROUTING") Then
				
				ElseIf instr(1,sFailReason,"SMTPSEND") Then							
				
				Else
					'Need RegEx to clean up foreign server responses to 5XX a.b.c with actual values returned
					strFailReasonTemp = ExtractResponseCode(strFailReason)
					
					If len(strFailReasonTemp) > 1 Then
						If aFileLine(iSource) = "ROUTING" and aFileLine(iSourceContext) = "" Then
							'suppress duplicate FAIL event counting
							Exit Function
						ElseIf aFileLine(iSource) = "SMTP" and aFileLine(iConnectorId)<>"" Then
							strFailReason = strFailReasonTemp & " " & aFileLine(iSource) & "(" & aFileLine(iConnectorId) & ")"								
						Else
							strFailReason = strFailReasonTemp & " " & aFileLine(iSource) & "(" & aFileLine(iSourceContext) & ")"
						End If
					Else
						'Wscript.echo vbcrlf & strFailReasonTemp
						strFailReason = "5XX a.b.c " & aFileLine(iSource) & "(" & aFileLine(iSourceContext) & ")"
					End If
					
					
				End If					
				
				strFailReason = replace(strFailReason,"'","")
				strFailReason = replace(strFailReason,"""","")
														
			Else 'Blank Recipient Status Found
				aSource = split(aFileLine(iSource),";")
				aSourceContext = split(aFileLine(iSourceContext),";")
				'Wscript.echo "sFailReason is blank, source = " & aSource(0) & "(" & aSourceContext(0) & ")"
				strFailReason = "5XX a.b.c " & aSource(0) & "(" & aSourceContext(0) & ")"
				
			End If
							
			If oFailReasonDictionary.Exists(strFailReason) Then
				iFailReasonCount = int(oFailReasonDictionary.Item(strFailReason))
				oFailReasonDictionary.Item(strFailReason) = iFailReasonCount + int(aFileLine(iRecipCount))
			Else
				oFailReasonDictionary.ADD strFailReason, int(aFileLine(iRecipCount))
			End If
							
			If aFileLine(iSource) = "AGENT" Then
				sFailSource = aFileLine(iSource) & ":" & aFileLine(iSourceContext)
			ElseIf aFileLine(iSource) = "SMTP" and aFileLine(iConnectorId)<>"" Then
				sFailSource = aFileLine(iSource) & ":" & """" & aFileLine(iConnectorId) & """"
			Else
				sFailSource = aFileLine(iSource)
			End If
			
			If aFileLine(iReturnPath) = "<>" Then
				'FAIL for null reverse path
				If oNullReturnFailSourceDictionary.Exists(sFailSource) Then
					iFailSourceCount = oNullReturnFailSourceDictionary.Item(sFailSource)
					oNullReturnFailSourceDictionary.Item(sFailSource) = iFailSourceCount + 1
				Else
					oNullReturnFailSourceDictionary.Add sFailSource, 1
				End IF
			Else
				'FAIL for valid reverse path
				iValidReturnFailCount = iValidReturnFailCount + 1
				If instr(1,lcase(aFileLine(iSender)),sInternalDNSSuffix) > 1 Then
					iValidReturnFailInternalSenderCount = iValidReturnFailInternalSenderCount + 1
				Else
					iValidReturnFailExternalSenderCount = iValidReturnFailExternalSenderCount + 1
				End If
			End If
			
			If oFailSourceDictionary.Exists(sFailSource) Then
				iFailSourceCount = oFailSourceDictionary.Item(sFailSource)
				oFailSourceDictionary.Item(sFailSource) = iFailSourceCount + 1
			Else
				oFailSourceDictionary.Add sFailSource, 1
			End IF
			
			If instr(1,lcase(aFileLine(iSender)),sInternalDNSSuffix) > 1 Then
				iFailInternalSenderCount = iFailInternalSenderCount + 1
			Else
				iFailExternalSenderCount = iFailExternalSenderCount + 1
			End If
			
			If instr(1,aFileLine(iReference),";") Then 
				aReference=split(aFileLine(iReference),";")				
				sReference = aReference(0)
			Else
				sReference=aFileLine(iReference)
			End If
			
			If sReference<>"" AND oNdrMsgIdDictionary.Exists(sReference) Then
				'Corresponding DSN found
				sEventDateTime = GetDateTime(aFileLine(iEventTime))
				'Wscript.echo sEventDateTime & "," & aFileLine(iClientName) & "," & strFailReason & "," & aFileLine(iReturnPath) & "," & sReference & "," & aFileLine(iRecipCount)
				oDsnFailureResults.writeline sEventDateTime & "," & aFileLine(iClientName) & "," & strFailReason & "," & aFileLine(iReturnPath) & "," & sReference & "," & aFileLine(iRecipCount) & "," & aFileLine(iRecipients)
				oNdrMsgIdDictionary.Remove(sReference)
				iFailAndNdrGenerated = iFailAndNdrGenerated + 1
			ElseIf aFileLine(iReturnPath)="<>" Then
				'Null ReturnPath (BADMAIL DSN)
			Else
				'Corresponding DSN not found for valid return path
				'Wscript.echo "WARNING: Could not find DSN for FAIL event for " & sFileLine
				sEventDateTime = GetDateTime(aFileLine(iEventTime))
				If bWriteDsnFailureNotFoundResults Then
					oDsnFailureNotFoundResults.writeline sEventDateTime & "," & aFileLine(iClientName) & "," & strFailReason & "," & aFileLine(iReturnPath) & "," & aFileLine(iMessageId) & "," & aFileLine(iRecipCount) & "," & aFileLine(iRecipients) & "," & aFileLine(iReference)
				End If
			End IF

			
				
		Case "DSN"			
			iDsnCount = iDsnCount + 1

			If oDsnContextDictionary.Exists(aFileLine(iSourceContext)) Then
				iDsnContextCount = oDsnContextDictionary.Item(aFileLine(iSourceContext))
				oDsnContextDictionary.Item(aFileLine(iSourceContext)) = iDsnContextCount + 1
			Else
				oDsnContextDictionary.Add aFileLine(iSourceContext),1 
			End If
			
			If aFileLine(iSourceContext) = "Failure" Then			
				iFileDsnFailureCount = iFileDsnFailureCount + 1				
				If oNdrMsgIdDictionary.Exists(aFileLine(iMessageId)) Then
					LogText "ERROR: More than one DSN event found for " & aFileLine(iMessageId)
				Else
					oNdrMsgIdDictionary.Add aFileLine(iMessageId),aFileLine
				End If
			End If

		
		Case "RESOLVE"
			iResolveCount = iResolveCount + 1
		
		Case "TRANSFER"
			iTransferCount = iTransferCount + 1
			iFileTransferCount = iFileTransferCount + 1			
			If oTransferContextDictionary.Exists(aFileLine(iSourceContext)) Then
				aExistTransferContext = oTransferContextDictionary.Item(aFileLine(iSourceContext))
				aExistTransferContext(0) = aExistTransferContext(0) + 1
				aExistTransferContext(1) = aExistTransferContext(1) + iPreTransferLatency
				oTransferContextDictionary.Item(aFileLine(iSourceContext)) = aExistTransferContext
			Else
				Redim aTransferContext(1)
				aTransferContext(0) = 1
				aTransferContext(1) = iPreTransferLatency
				oTransferContextDictionary.Add aFileLine(iSourceContext),aTransferContext
			End If
									
		Case "DUPLICATEDELIVER"
			sLocalServerName = aFileLine(iClientName)
			iDuplicateDeliver = iDuplicateDeliver + 1
			iDuplicateDeliverRecipCount = iDuplicateDeliverRecipCount + int(aFileLine(iRecipCount))
			If aFileLine(iRecipCount) = 1 Then
				aRecipients = Array(aFileLine(iRecipients))
			Else
				aRecipients = split(aFileLine(iRecipients),";")
			End If
			
			If TRUE then
			End IF
			
			For i = 0 to ubound(aRecipients)
				'Create dictionary to count top messages and recipients of duplicate deliveries
				sDuplicateDeliveryDictionaryKey = aFileLine(iSender) & "," & aFileLine(iMessageid) & "," & aRecipients(i)
				If oDuplicateDeliveryDictionary.Exists(sDuplicateDeliveryDictionaryKey) Then
					iDuplicateDeliveryDictionaryItem = int(oDuplicateDeliveryDictionary.Item(sDuplicateDeliveryDictionaryKey))
					oDuplicateDeliveryDictionary.Item(sDuplicateDeliveryDictionaryKey) = iDuplicateDeliveryDictionaryItem + 1
				Else
					oDuplicateDeliveryDictionary.Add sDuplicateDeliveryDictionaryKey, 1
				End If
			Next
			
		
		Case "REDIRECT"
			iRedirectCount = iRedirectCount + 1
			
		Case "BADMAIL" 
			iBadmailCount = iBadmailCount + 1
		
		Case "DEFER"
			iDeferCount = iDeferCount + 1
			iFileDeferCount = iFileDeferCount + 1
			strDeferSource = aFileLine(iSource) & "(" & aFileLine(iSourceContext) & ")"
			If aFileLine(iMessageInfo) = "0001-01-01T08:00:00.000Z" Then
				iDeferDurationSec = 0
			Else
				iDeferDurationSec = datediff("s",GetDateTime(aFileLine(iEventTime)),GetDateTime(aFileLine(iMessageInfo)))
			End If
			'Wscript.echo aFileLine(iMessageId) & vbtab & "DEFER(" & aFileLine(iSourceContext) & ") duration = " & iDeferDurationSec
			If oDeferContextDictionary.Exists(strDeferSource) Then
				aExistDeferContext = oDeferContextDictionary.Item(strDeferSource)
				aExistDeferContext(0) = aExistDeferContext(0) + 1
				aExistDeferContext(1) = aExistDeferContext(1) + iDeferDurationSec
				oDeferContextDictionary.Item(strDeferSource) = aExistDeferContext
			Else
				Redim aDeferContext(1)
				aDeferContext(0) = 1
				aDeferContext(1) = iDeferDurationSec
				oDeferContextDictionary.Add strDeferSource,aDeferContext
			End If
			
		Case "POISONMESSAGE"
			sLocalServerName = sLastLocalServerName
			iPoisonCount = iPoisonCount + 1
			If instr(1,aFileLine(iMessageId),sInternalDNSSuffix) Then
				'Internal MessageId found
				iInternalPoisonCount = iInternalPoisonCount  +1
			Else
				'External MessageId found
			End IF
			
		Case "SUBMIT"
			iSubmitEventCount = iSubmitEventCount + 1
			aSourceContext = split(replace(aFileLine(iSourceContext),"""",""),",")
			For i = 0 to ubound(aSourceContext)
				aSourceContextEntry = split(aSourceContext(i),":")
				Select Case trim(aSourceContextEntry(0))
					Case "MDB"						
					Case "Mailbox"
					Case "Event"
					Case "MessageClass"
						strMessageClass = aSourceContextEntry(1)
					Case "CreationTime"
					Case "ClientType"
						strClientType = aSourceContextEntry(1)
					Case Else
				End Select
			Next
			strSubmitDictionaryKey = aFileLine(iClientName) & "," & aFileLine(iServerName) & "," & strClientType & "," & strMessageClass
			If oSubmitDictionary.Exists(strSubmitDictionaryKey) Then
				oSubmitDictionary.Item(strSubmitDictionaryKey) = oSubmitDictionary.Item(strSubmitDictionaryKey) + 1
			Else
				oSubmitDictionary.Add strSubmitDictionaryKey, 1
			End If
				
		
		Case "RESUBMIT"
			iReSubmitEventCount = iReSubmitEventCount + 1
			sCurrentEventTime = GetDateTime(aFileLine(iEventTime))
			
			sReSubmitDictionaryKey = aFileLine(iServername) & ":" & aFileLine(iSource) & "(" & aFileLine(iSourceContext) & ")"
			If oReSubmitDictionary.Exists(sReSubmitDictionaryKey) Then
				aReSubmitDictionaryItem = oReSubmitDictionary.Item(sReSubmitDictionaryKey)
				aReSubmitDictionaryItem(0) = aReSubmitDictionaryItem(0) + 1
				aReSubmitDictionaryItem(2) = sCurrentEventTime
				aReSubmitDictionaryItem(3) = aReSubmitDictionaryItem(3) + int(aFileLine(iTotalBytes))
				oReSubmitDictionary.Item(sReSubmitDictionaryKey) = aReSubmitDictionaryItem
			Else			
				oReSubmitDictionary.add sReSubmitDictionaryKey, array(1,sCurrentEventTime,sCurrentEventTime,int(aFileLine(iTotalBytes)))
			End If
			
			sReSubmitMessageDictionaryKey = aFileLine(iServername) & ":" & aFileLine(iSource) & ":" & aFileLine(iSourceContext) & ":" & aFileLine(iMessageid) & ":" & aFileLine(iRecipients) 
			If oReSubmitMessageDictionary.Exists(sReSubmitMessageDictionaryKey) and aFileLine(iMessageid) <> "" Then
				LogText vbcrlf & "**Warning: found multiple resubmissions for same message"
				LogText "sReSubmitMessageDictionaryKey = " & sReSubmitMessageDictionaryKey
				LogText "ORIGINAL EVENT: " & oReSubmitMessageDictionary.Item(sReSubmitMessageDictionaryKey)
				LogText "CURRENT EVENT: " & sFileLine
			Else
				oReSubmitMessageDictionary.Add sReSubmitMessageDictionaryKey,sFileLine
				'LogText "oReSubmitMessageDictionary.add " & sReSubmitMessageDictionaryKey
			End If			
		
		Case Else
				LogText sFileLine
				
			
	    
	End Select

	If sLocalServerName = "" Then
		LogText vbcrlf & "**Warning: Found Event with server name" & vbCRLF & sFileLine & vbCRLF
	End If			
	
	If sLocalServerName <> sLastLocalServerName AND iLogFileEventCount > 1 Then
		LogText vbcrlf & "**Warning: Found Event with different server name " & sLocalServerName & " which does not match " & sLastLocalServerName
	End If
	
	If aFileLine(iSource) = "SMTP" AND aFileLine(iConnectorId) <> "" Then
		strMySource = "SMTP(" & replace(aFileLine(iConnectorId),":","_") & ")"
	ElseIf aFileLine(iEventID) = "TRANSFER" OR aFileLine(iSource) = "AGENT" OR aFileLine(iEventID) = "DSN" Then
		strMySource = aFileLine(iSource) & "(" & replace(aFileLine(iSourceContext),":","_") & ")"
	Else
		strMySource = aFileLine(iSource)
	End If
	
	sEventTimeDictionaryKey = sLocalServerName & ":" & aFileLine(iEventID) & ":" & strMySource & ":" & month(sCurrentEventTime) & ":" & day(sCurrentEventTime) & ":" & hour(sCurrentEventTime)
	
	If oEventTimeDictionary.Exists(sEventTimeDictionaryKey) Then
		oEventTimeDictionary.Item(sEventTimeDictionaryKey) = oEventTimeDictionary.Item(sEventTimeDictionaryKey) + 1				
	Else
		oEventTimeDictionary.Add sEventTimeDictionaryKey, 1
	End If	
	  

End Function


Function ParseCsvLine(byVal sString, ByRef aString)
	On Error Resume Next
	'bDebugOut = TRUE
	iEnterParseCsvLine = iEnterParseCsvLine + 1
	'Input=CSV,Output=Array
	i = 0
	If bDebugOut Then LogText vbcrlf & vbcrlf & sString & vbcrlf
	Dim regEx, Match, Matches   ' Create variable.
	Set regEx = New RegExp   ' Create a regular expression.
	regEx.Pattern = ",(?=([^""]*""[^""]*"")*(?![^""]*""))"
	regEx.IgnoreCase = True   ' Set case insensitivity.
	regEx.Global = True   ' Set global applicability.
	Set Matches = regEx.Execute(sString)
	iIndex = 1
	If bDebugOut Then LogText "Values found: " & Matches.count
	ReDim Preserve aString(i)
	For Each Match in Matches   ' Iterate Matches collection.
		aString(i) = mid(sString,iIndex,Match.FirstIndex-iIndex+1)
		If bDebugOut Then LogText i & vbtab & aString(i)
		iIndex = Match.FirstIndex+2
		i = i + 1	
		ReDim Preserve aString(i)
	Next
	If len(sString)>iIndex Then
		aString(i) = right(sString,len(sString)-iIndex+1)
	End If
	If bDebugOut Then LogText i & vbtab & aString(i)
	
	If bDebugOut Then If iEnterParseCsvLine=3 Then Exit Function
	'bDebugOut = FALSE
	
End Function

Function ExtractResponseCode(byVal sInput)
	On Error Resume Next
	'Input=RecipientResponse,Output=ResponseCode
	sOutput = ""
	'Wscript.echo vbcrlf & vbcrlf & "sInput = " & sInput & vbcrlf
	Dim regEx, Match, Matches   ' Create variable.
	Set regEx = New RegExp   ' Create a regular expression.
	regEx.Pattern = "(5\d{2}) (\d\.\d\.\d)"
	regEx.IgnoreCase = True   ' Set case insensitivity.
	regEx.Global = True   ' Set global applicability.
	Set oMatches = regEx.Execute(sInput)
	For each oMatch in oMatches
		'Wscript.echo "oMatch = " & oMatch.Value
		sOutput = sOutput & " " & oMatch.Value
		Exit For
	Next
	If sOutput="" Then
		regEx.Pattern = "(5\d{2}) "
		regEx.IgnoreCase = True   ' Set case insensitivity.
		regEx.Global = True   ' Set global applicability.
		Set oMatches = regEx.Execute(sInput)
		For each oMatch in oMatches
			'Wscript.echo "oMatch = " & oMatch.Value
			sOutput = sOutput & " " & oMatch.Value
			Exit For
		Next
		sOutput = sOutput & "a.b.c"
	End If
	ExtractResponseCode = trim(sOutput)
End Function

Function FixResponseCode(byVal sInput)
	On Error Resume Next
	'Input=RecipientResponse,Output=ResponseCode
	sOutput = sInput
	'Wscript.echo vbcrlf & vbcrlf & "sInput = " & sInput & vbcrlf
	Dim regEx, Match, Matches   ' Create variable.
	Set regEx = New RegExp   ' Create a regular expression.
	regEx.Pattern = "\d\.\d\.\d ([M2MCVT|PICKUP|QUEUE|REPLAY|RESOLVER|ROUTING|SMTPSEND|NONSMTPGW|STOREDRV].*?'*;+)"
	regEx.IgnoreCase = True   ' Set case insensitivity.
	regEx.Global = True   ' Set global applicability.
	Set oMatches = regEx.Execute(sInput)
	For each oMatch in oMatches
		'If instr(1,sInput,"RESOLVER.RST") Then Wscript.echo "oMatch = " & oMatch.Value
		sTempExp = replace(oMatch.Value,";",":")
		sOutput = replace(sOutput,oMatch.Value,sTempExp)
		'Wscript.echo "New Response = " & sOutput	
	Next
	FixResponseCode = sOutput
	'If instr(1,sOutput,"RESOLVER.RST") Then Wscript.echo "ExtractResponseCode = " & trim(sOutput)
End Function




Function GetDateTime(byVal sLogDateTimeTFormat)
	On Error Resume Next
	'bDebugOut = TRUE
	If sLogDateTimeTFormat<>"" Then
		If bDebugOut Then LogText "sLogDateTimeTFormat=" & sLogDateTimeTFormat
		strDate = left(sLogDateTimeTFormat,instr(1,sLogDateTimeTFormat,"T")-1)
		If bDebugOut Then LogText strDate
		MyDate = DateValue(strDate)
		strTime = mid(sLogDateTimeTFormat,instr(1,sLogDateTimeTFormat,"T")+1,len(sLogDateTimeTFormat)-instr(1,sLogDateTimeTFormat,"T")-5)
		If bDebugOut Then LogText strTime
		MyTime = TimeValue(strTime)
		GetDateTime = MyDate & " " & MyTime				
	Else
		GetDateTime = ""
	End If
	bDebugOut = FALSE
End Function

Sub WriteSummary

	On Error Resume Next
	
	oSummaryResults.WriteLine "Script Version " & strVersion
	oSummaryResults.WriteLine GetProcessInfo(FALSE)
	
	oSummaryResults.WriteLine vbcrlf

	oSummaryResults.WriteLine "Oldest Log File Processed: " & sOldestLogDate & " UTC"
	oSummaryResults.WriteLine "Newest Log File Processed: " & sNewestLogDate & " UTC"
	oSummaryResults.WriteLine "Total Log Files Processed: " & iEnterParseFileFunctionTotal
	oSummaryResults.WriteLine "Total Log Files Processed Out of Sequence: " & iTotalFilesProcessOutOfSequenceCount
	oSummaryResults.WriteLine "Total Log Files Parsed: " & iTotalFilesParsed
	oSummaryResults.WriteLine "Total Log Files Parsed Out of Sequence: " & iTotalFilesParsedOutOfSequenceCount 
	

	If iTotalFilesParsed > 0 Then oSummaryResults.WriteLine "Total Log Duration (Min): " & round(iTotalFileDateTimeDiffMin,1) & " (" & round(iTotalFileDateTimeDiffMin/iTotalFilesParsed,1) & " min/log)"
	oSummaryResults.WriteLine "Maximum File Duration: " & iFileDurationMax & " min"
	oSummaryResults.WriteLine "Minimum File Duration: " & iFileDurationMin & " min"
	oSummaryResults.WriteLine "Log File Statistics: " & sLogStatistics
	oSummaryResults.WriteLine "Valid FAIL events with no DSN Failures found, see " & sDsnFailureNotFoundResults & " for details)"
	oSummaryResults.WriteLine vbCRLF
	If iTotalFilesParsed > 0 Then oSummaryResults.WriteLine "Total Events processed: " & iEventCount & " (Avg Events/Log: " & (iEventCount/iTotalFilesParsed) & ")"
	If iMsgIdCount > 0 Then oSummaryResults.WriteLine "Total Message Id's Processed: " & iMsgIdCount & " (Avg Events/MsgId: " & (iEventCount/iMsgIdCount) & ", Max: " & iMaxLogEventPerMsgId & ", Min: " & iMinLogEventPerMsgId & ")"
	If iReceiveCount > 0 Then  oSummaryResults.WriteLine "Total Messages received: " & iReceiveCount & " (Avg Msg Size: " & round((iTotalMsgBytes/iReceiveCount)/1024,0) & " KB)"
	oSummaryResults.WriteLine "Total Messages sent: " & iSendCount
	oSummaryResults.WriteLine "Total Messages delivered: " & iDeliverCount
	oSummaryResults.WriteLine "Total Messages delivered (duplicate): " & iDuplicateDeliver & " (" & iDuplicateDeliverRecipCount & " recipients)"
	oSummaryResults.WriteLine "Total Resolve: " & iResolveCount
	If iTransferCount > 0 Then oSummaryResults.WriteLine "Total Transfer: " & iTransferCount & " (" & round(100*iPreTransferLatencyGtSLA/iTransferCount,1) & "% exceeded " & iPreTransferLatencySLA & " sec before transfer, average " & round(iTotalPreTransferLatency/iTransferCount,1) & " sec latency before transfer)"
	oSummaryResults.WriteLine "Total Expand: " & iExpandCount
	If iFailCount > 0 Then oSummaryResults.WriteLine "Total Fail: " & iFailCount & " (" & round(iFailInternalSenderCount/iFailCount*100,1) & "% *." & sInternalDNSSuffix & " senders, " & round(iFailExternalSenderCount/iFailCount*100,1) & "% external senders)"
	oSummaryResults.WriteLine "Total Fail with NDR: " & iFailAndNdrGenerated
	If iValidReturnFailCount > 0 Then oSummaryResults.WriteLine "Total Fail (valid Return Path): " & iValidReturnFailCount & " (" & round(iValidReturnFailInternalSenderCount/iValidReturnFailCount*100,1) & "% *." & sInternalDNSSuffix & " senders, " & round(iValidReturnFailExternalSenderCount/iValidReturnFailCount*100,1) & "% external senders)"
	oSummaryResults.WriteLine "Total Fail (Recipient): " & iFailRecipientCount
	If iTotalFileDateTimeDiffMin Then oSummaryResults.WriteLine "Avg Fail Events/Min: " & round(iFailCount/iTotalFileDateTimeDiffMin,1)
	oSummaryResults.WriteLine "Max Fail Events/Min: " & iMaxAvgFailPerMin
	oSummaryResults.WriteLine "Max DSN(Fail) Events/Min: " & iMaxAvgFileDsnCount
	oSummaryResults.WriteLine "Total DSN Generated: " & iDsnCount
	oSummaryResults.WriteLine "Total DSN Badmail: " & iBadMailCount 
	If (iDsnCount+iBadMailCount) > 0 Then oSummaryResults.WriteLine "Total DSN: " & (iDsnCount+iBadMailCount) & " (Percent DSN Badmail: " & round((iBadmailCount/(iDsnCount+iBadMailCount))*100,1) & "%)"
	If iDeliverCount > 0 Then
		oSummaryResults.WriteLine "Total DSN Delivered from Internal domains: " & iMERDeliverCount
		oSummaryResults.WriteLine "Total Null Reverse Path Delivered: " & iInternalNullReversePathDeliverCount+iExternalNullReversePathDeliverCount
		oSummaryResults.WriteLine "Total Null Reverse Path Delivered from Internal: " & iInternalNullReversePathDeliverCount & " (" & round(iInternalNullReversePathDeliverCount/(iInternalNullReversePathDeliverCount+iExternalNullReversePathDeliverCount)*100,1) & "%)"
		oSummaryResults.WriteLine "Total Null Reverse Path Delivered from External: " & iExternalNullReversePathDeliverCount& " (" & round(iExternalNullReversePathDeliverCount/(iInternalNullReversePathDeliverCount+iExternalNullReversePathDeliverCount)*100,1) & "%)"
	End If
	oSummaryResults.WriteLine "Total Defer: " & iDeferCount
	If iPoisonCount>0 Then
		oSummaryResults.WriteLine "Total Poison: " & iPoisonCount & " (" & round(iInternalPoisonCount/iPoisonCount,1) & "% internal)"
	Else
		oSummaryResults.WriteLine "Total Poison: " & iPoisonCount
	End If
	oSummaryResults.WriteLine "Total Split Errors: " & iSplitError

	iSumEvent = iReceiveCount + iSendCount + iDeliverCount + iDuplicateDeliver + iDsnCount + iResolveCount + iTransferCount + iExpandCount + iFailCount + iRedirectCount + iBadmailCount + iSplitError + iDeferCount + iSubmitEventCount

	If (iEventCount - iSumEvent) > 4 Then
		oSummaryResults.WriteLine "iSumEvent: " & iSumEvent
	End If

	oSummaryResults.WriteLine vbcrlf
	oSummaryResults.WriteLine "Total MailSubmission Submit Events: " & iSubmitEventCount

	oSummaryResults.WriteLine vbcrlf
	If iReceiveSmtpCount > 0 Then
		oSummaryResults.WriteLine "Total SMTP Receive: " & iReceiveSmtpCount & " (RecipCount: " & iReceiveSmtpRecipCount & ", recipients/receive: " & round(iReceiveSmtpRecipCount/iReceiveSmtpCount,1) & ")"
		a = oReceiveClientDictionary.Keys
		For i = 0 to oReceiveClientDictionary.Count - 1
			aReceiveClientDictionaryItem = oReceiveClientDictionary.Item(a(i))
			oReceiveClientFileOut.writeline replace(a(i),":",",") & "," & aReceiveClientDictionaryItem(0) & "," & round(aReceiveClientDictionaryItem(1)/aReceiveClientDictionaryItem(0),0)
			'oSummaryResults.WriteLine a(i) & vbtab & oReceiveClientDictionary.Item(a(i))
		Next
	End If
	
	If iReceiveStoreDriverCount > 0 Then
		oSummaryResults.WriteLine "Total StoreDriver Receive: " & iReceiveStoreDriverCount & " (RecipCount: " & iReceiveSDRecipCount & ", Recipients/submit: " & round(iReceiveSDRecipCount/iReceiveStoreDriverCount,1) & ", Unique senders: " & oTopSendersbySubmitDictionary.Count & ", Submissions/sender: " & round(iReceiveStoreDriverCount/oTopSendersbySubmitDictionary.Count,1) & ")"
	End If

	If iSendCount > 0 Then
		oSummaryResults.WriteLine "Total SMTP Send: " & iSendCount + iDeliverCount & " (RecipCount: " & iSendSmtpRecipCount & ", recipients/send: " & round(iSendSmtpRecipCount/iSendCount,1) & ")"
	End If
	If iDeliverCount > 0 Then
		oSummaryResults.WriteLine "Total StoreDriver Deliver: " & iDeliverCount & " (RecipCount: " & iSdDeliverRecipCount & ", Recipients/deliver: " & round(iSdDeliverRecipCount/iDeliverCount,1) & ", Unique Recipient Mailboxes: " & oTopRecipientsDictionary.Count & ", Deliveries/Mailbox: " & round(iSdDeliverRecipCount/oTopRecipientsDictionary.Count,1) & ")"
	End If
	oSummaryResults.WriteLine "For details on SMTP Send and StoreDriver Deliver NextHop statistics, see " & sNextHopResults
	a = oNextHopServerDictionary.Keys
	For i = 0 to oNextHopServerDictionary.Count - 1
		aNextHopServerDictionaryItem = oNextHopServerDictionary.Item(a(i))
		oNextHopResults.writeline replace(a(i),":",",") & "," & aNextHopServerDictionaryItem(0) & "," & round(aNextHopServerDictionaryItem(1)/aNextHopServerDictionaryItem(0),0) & "," & aNextHopServerDictionaryItem(2) & "," & round(100*aNextHopServerDictionaryItem(2)/aNextHopServerDictionaryItem(0),2) & "," & aNextHopServerDictionaryItem(3) & "," & round(100*aNextHopServerDictionaryItem(3)/aNextHopServerDictionaryItem(0),2)
	Next
	
	
	If iExpandCount > 0 Then
	oSummaryResults.WriteLine "Total Expand: " & iExpandCount & " (RecipCount: " & iExpandRecipCount & ", Recipients/expand: " & round(iExpandRecipCount/iExpandCount,1) & ", Unique DL's: " & iUniqueDL & ", Max RecipCount: " & iMaxExpandRecipCount & ")"
	End IF

	a = oExpandedDLDictionary.Keys
	For i = 0 to oExpandedDLDictionary.Count - 1
		aExpandDetail = oExpandedDLDictionary.Item(a(i))
		'0=RecipCount,1=ExpandCount,2=TotalExpandLatency
		oExpandFileOut.writeline i & "," & a(i) & "," & aExpandDetail(0) & "," & aExpandDetail(1) & "," & round(aExpandDetail(2)/aExpandDetail(1),3)
	Next

	oSummaryResults.WriteLine vbcrlf & "Unique FAIL Recipient Status codes: " & vbcrlf 

	a = oFailReasonDictionary.Keys
	For i = 0 to oFailReasonDictionary.Count - 1
			
		If a(i) = "550 5.1.1 RESOLVER.ADR.RecipNotFound" Then
			oSummaryResults.WriteLine i & "," & a(i) & "," & oFailReasonDictionary.Item(a(i)) & " (Unique Recipients: " & oRecipNotFoundDictionary.Count & ")"
		ElseIf a(i) = "550 5.2.2 STOREDRV.Deliver: mailbox full" Then
			oSummaryResults.WriteLine i & "," & a(i) & "," & oFailReasonDictionary.Item(a(i)) & " (Unique Recipients: " & oMailboxFullDictionary.Count & ")"
		Else
			oSummaryResults.WriteLine i & "," & a(i) & "," & oFailReasonDictionary.Item(a(i)) 
		End If
		
	Next
	
	a = oReSubmitDictionary.Keys
	If oReSubmitDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbcrlf & "Resubmit events found:"
		oSummaryResults.WriteLine vbcrlf & "ServerName:Source(DumpsterDn)" & vbtab & "Count" & vbtab & "TotalBytes" & vbtab & "BeginTimeStamp" & vbtab & "Duration(sec)"
		For i = 0 to oReSubmitDictionary.Count - 1
			aReSubmitDictionaryItem = oReSubmitDictionary.Item(a(i))			
			oSummaryResults.WriteLine a(i) & vbtab & aReSubmitDictionaryItem(0) & vbtab & aReSubmitDictionaryItem(3) & vbtab & aReSubmitDictionaryItem(1) & " UTC" & vbtab & datediff("s",aReSubmitDictionaryItem(1),aReSubmitDictionaryItem(2))
		Next
	End If	
	
	oSummaryResults.WriteLine vbcrlf & "Delivery Latency exceeded " & iDeliveryLatencySLAExceptionLoggingThreshold & " sec for " & oDeliveryLatencyExceptionDictionary.Count & " messages" 
	oSummaryResults.WriteLine "Delivery Latency Exception Details at " & sDeliveryLatencyExceptionResults
	a = oDeliveryLatencyExceptionDictionary.Keys
	For i = 0 to oDeliveryLatencyExceptionDictionary.Count - 1
		'iMessageInfo,iClientName,iTotalBytes,iSender,iRecipCount,iDeliverCount,iMinDeliveryLatency,iMaxDeliveryLatency
		aExistDeliveryLatencyExceptionDictionaryItem = oDeliveryLatencyExceptionDictionary.Item(a(i))
		sDeliveryLatencyExceptionResultsLine = a(i)
		for j = 0 to 7 
			sDeliveryLatencyExceptionResultsLine = sDeliveryLatencyExceptionResultsLine & "," & aExistDeliveryLatencyExceptionDictionaryItem(j)
		next
		oDeliveryLatencyExceptionResults.writeline sDeliveryLatencyExceptionResultsLine
	Next

	oSummaryResults.WriteLine vbcrlf & "DSN Failure Details at " & sDsnFailureResults
	
	oSummaryResults.WriteLine vbcrlf & "Mailbox Full Recipient Details at " & sMbxFullRecipResults
	a = oMailboxFullDictionary.Keys
	For i = 0 to oMailboxFullDictionary.Count - 1
		oMbxFullRecipResults.writeline a(i) & "," & oMailboxFullDictionary.Item(a(i))
	Next
	

	oSummaryResults.WriteLine vbcrlf & "FAIL Event Sources: " & vbcrlf 

	a = oFailSourceDictionary.Keys
	For i = 0 to oFailSourceDictionary.Count - 1
		iValidDSN = 0
		iValidDSN = oFailSourceDictionary.Item(a(i)) - oNullReturnFailSourceDictionary.Item(a(i))
		oSummaryResults.WriteLine i & "," & a(i) & "," & oFailSourceDictionary.Item(a(i)) & " (Null ReturnPath: " & oNullReturnFailSourceDictionary.Item(a(i)) & ", Valid ReturnPath: " & iValidDSN & ")"
	Next


	a = oTransferContextDictionary.Keys
	If oTransferContextDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbcrlf & "Transfer Source Context: " & vbcrlf 
		oSummaryResults.WriteLine "Id,TransferContext,MsgCount,AvgLatencySec"
		For i = 0 to oTransferContextDictionary.Count - 1
			aExistTransferContext = oTransferContextDictionary.Item(a(i)) 
			oSummaryResults.WriteLine i & "," & a(i) & "," & aExistTransferContext(0) & "," & round(aExistTransferContext(1)/aExistTransferContext(0),1)
		Next
	End If

	a = oSDSubmitDomainDictionary.Keys
	oSummaryResults.WriteLine vbcrlf & "Unique StoreDriver Sender Domains: " & oSDSubmitDomainDictionary.Count & vbcrlf 
	For i = 0 to oSDSubmitDomainDictionary.Count - 1
		if oSDSubmitDomainDictionary.Item(a(i)) > 5 Then oSummaryResults.WriteLine i & "," & a(i) & "," & oSDSubmitDomainDictionary.Item(a(i)) 
	Next

	a = oDeferContextDictionary.Keys
	If oDeferContextDictionary.Count > 0 Then 
		oSummaryResults.WriteLine vbcrlf & "Unique Defer Source Context: " & oDeferContextDictionary.Count & vbcrlf 
		oSummaryResults.WriteLine "Id,Context,MsgCount,AvgLatencySec"
		For i = 0 to oDeferContextDictionary.Count - 1
			aExistDeferContext = oDeferContextDictionary.Item(a(i))
			oSummaryResults.WriteLine i & "," & a(i) & "," & aExistDeferContext(0) & "," & round(aExistDeferContext(1)/aExistDeferContext(0),1)
		Next
	End If

	a = oDsnContextDictionary.Keys
	oSummaryResults.WriteLine vbcrlf & "Unique DSN Source Context: " & oDsnContextDictionary.Count & vbcrlf 
	For i = 0 to oDsnContextDictionary.Count - 1
		oSummaryResults.WriteLine i & "," & a(i) & "," & oDsnContextDictionary.Item(a(i)) 
	Next

	a = oIMCEADictionary.Keys
	oSummaryResults.WriteLine vbcrlf & "Unique Encapsulated Addresses: " & oIMCEADictionary.Count & vbcrlf 
	For i = 0 to oIMCEADictionary.Count - 1
		oSummaryResults.WriteLine i & "," & a(i) & "," & oIMCEADictionary.Item(a(i)) 
	Next
	
	If oPFSubjectDictionary.Count > 0 Then
		a = oPFSubjectDictionary.Keys
		oSummaryResults.WriteLine vbcrlf & "Public Folder Message Types: " & oPFSubjectDictionary.Count & vbcrlf 
		For i = 0 to oPFSubjectDictionary.Count - 1
			oSummaryResults.WriteLine i & "," & a(i) & "," & oPFSubjectDictionary.Item(a(i)) 
		Next
	Else
		oSummaryResults.WriteLine vbcrlf & "Public Folder Message Types: 0" & vbcrlf 
	End If
	
	If oRecipNotFoundDictionary.Count > 0 Then	
		Set oRecipientNotFoundResults = fso.OpenTextFile(sRecipientNotFoundResults,ForAppending,1)
		oRecipientNotFoundResults.writeline "RecipientNotFound,Count"
		iMyRecipNotFound = 0
		a = oRecipNotFoundDictionary.Keys
		For i = 0 to oRecipNotFoundDictionary.Count - 1
			oRecipientNotFoundResults.writeline a(i) & "," & oRecipNotFoundDictionary.Item(a(i))
		Next
	End If
	
	If oNdrMsgIdDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbCrlf & "DSN Failures without FAIL event found (" & oNdrMsgIdDictionary.Count & "):" 
		a = oNdrMsgIdDictionary.Keys
		For i = 0 to oNdrMsgIdDictionary.Count-1
			aFileLine = oNdrMsgIdDictionary.Item(a(i))
			oSummaryResults.WriteLine i & vbtab & aFileLine(iEventTime) & vbtab & aFileLine(iServerName) & vbtab & aFileLine(iMessageId) & vbtab & aFileLine(iRecipients)
		Next
	End If

	oSummaryResults.WriteLine vbCrlf & "Valid FAIL events with no DSN Failures found, see " & sDsnFailureNotFoundResults & " for details)"

	If oDomainExpiredDictionary.Count > 0 Then
		Set oDomainExpiredResults = fso.OpenTextFile(sDomainExpiredResults,ForAppending,1)
		If fso.FileExists(sBlockedDomainList) Then fso.DeleteFile(sBlockedDomainList)
		Set oBlockedDomainList = fso.OpenTextFile(sBlockedDomainList,ForAppending,1)
		oDomainExpiredResults.writeline "Domain,RecipFailures,MsgFailures,OriginalExpireDateTime,LastExpireDateTime,MTBF(Hours)"
		'0=RecipCount,1=MsgCount,2=OrigDateTime,3=TotalSec,4=LastDateTime
		oSummaryResults.WriteLine vbCrlf & oDomainExpiredDictionary.Count & " unique domains where encountered Queue.Expired, see " & sDomainExpiredResults & " for details" 
		a = oDomainExpiredDictionary.Keys
		For i = 0 to oDomainExpiredDictionary.Count-1
			aDomainExpiredDictionaryEntry = oDomainExpiredDictionary(a(i))
			err.Clear
			If aDomainExpiredDictionaryEntry(1) > 0 Then
				iMTBFSec = round(aDomainExpiredDictionaryEntry(3)/aDomainExpiredDictionaryEntry(1)/3600,0)
				If err.number <> 0 Then
					LogText now & vbtab & "ERROR" & vbtab & "0x" & hex(err.number) & ":" & err.Description & " when attempting '" & aDomainExpiredDictionaryEntry(3) & "/" & aDomainExpiredDictionaryEntry(1) & "'"
				End If
			Else
				iMTBFSec = 0
			End If
			oDomainExpiredResults.writeline a(i) & "," & aDomainExpiredDictionaryEntry(0) & "," & aDomainExpiredDictionaryEntry(1) & "," & aDomainExpiredDictionaryEntry(2) & "," & aDomainExpiredDictionaryEntry(4) & "," & iMTBFSec
			oBlockedDomainList.writeline a(i) & "," & aDomainExpiredDictionaryEntry(0) & "," & aDomainExpiredDictionaryEntry(1) & "," & aDomainExpiredDictionaryEntry(2) & "," & aDomainExpiredDictionaryEntry(3) & "," & aDomainExpiredDictionaryEntry(4)
		Next
		LogText now & vbtab & "INFO" & vbtab & "Done Saving Blocked Domain List Containing " & oDomainExpiredDictionary.Count & " Entries"
		oDomainExpiredResults.close
		oBlockedDomainList.close
	End If
	
	
	If oTopSendersbySubmitDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbCrlf & "Top Senders (by StoreDriver Receive, internal only) available in " & sTopSendersbySubmitResults & " (senders above average)"
		a = oTopSendersbySubmitDictionary.Keys
		For i = 0 to oTopSendersbySubmitDictionary.Count-1
			If oTopSendersbySubmitDictionary(a(i)) > (iReceiveStoreDriverCount / oTopSendersbySubmitDictionary.Count) Then
				oTopSendersbySubmitResults.writeline replace(a(i),":",",") & "," & oTopSendersbySubmitDictionary(a(i))
			End If
		Next
	End If
	
	If oTopSendersbyDeliverDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbCrlf & "Top Senders (by DELIVER including internal/external) available in " & sTopSendersbyDeliverResults & " (senders above average)"
		a = oTopSendersbyDeliverDictionary.Keys
		For i = 0 to oTopSendersbyDeliverDictionary.Count-1
			If oTopSendersbyDeliverDictionary(a(i)) > (iDeliverCount / oTopSendersbyDeliverDictionary.Count) Then
				oTopSendersbyDeliverResults.writeline a(i) & "," & oTopSendersbyDeliverDictionary(a(i))
			End If
		Next
	End If
	
	If oTopRecipientsDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbCrlf & "Top Recipients (internal only) available in " & sTopRecipientResults & " (recipients above average)"
		a = oTopRecipientsDictionary.Keys
		For i = 0 to oTopRecipientsDictionary.Count-1
			If oTopRecipientsDictionary(a(i)) > (iSdDeliverRecipCount / oTopRecipientsDictionary.Count) Then
				oTopRecipientResults.writeline replace(a(i),":",",") & "," & oTopRecipientsDictionary(a(i))
			End If
		Next
	End If
	
	If oEventTimeDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbCrlf & "Event Time Distribution available in " & sEventTimeDistribution
		Set oEventTimeDistribution = fso.OpenTextFile(sEventTimeDistribution,ForAppending,1)
		oEventTimeDistribution.writeline "Server,Event,Source,Month,Day,Hour,Count"
		a = oEventTimeDictionary.Keys
		For i = 0 to oEventTimeDictionary.Count-1
			oEventTimeDistribution.writeline replace(a(i),":",",") & "," & oEventTimeDictionary(a(i))
		Next
		oEventTimeDistribution.close
		Set oEventTimeDistribution = Nothing
	End If
	
	If oLatencyTrackerDictionary.Count > 0 Then
		Set oLatencyTrackerResults = fso.OpenTextFile(sLatencyTrackerResults,ForAppending,1)		
		oSummaryResults.WriteLine vbCrlf & "E14 Latency Tracker Info at " & sLatencyTrackerResults
		oSummaryResults.WriteLine "Individual Component Latency SLA Goal is " & iIndividualComponentLatencySLA & " seconds"
		oSummaryResults.WriteLine "Stage Component Latency SLA Goal is " & iStageComponentLatencySLA & " seconds"		
		a = oLatencyTrackerDictionary.Keys
		'0=MessageCountGt1Sec,1=TotalLatency,2=IndividualComponentSLA,3=StageComponentSLA,4=ServerSLA,5=TotalInvocationsGt1Sec,6=MaxInvocationCount,7=MaxLatency
		oLatencyTrackerResults.WriteLine "ServerFqdn,Component,AvgLatency,%Lt1Sec,%Lt" & iIndividualComponentLatencySLA & "Sec,%Lt" & iStageComponentLatencySLA & "Sec,%Lt" & iServerLatencySLA & "Sec,TotalInvocationsGt1Sec,MaxInvocationCount,MaxLatency"
		For i = 0 to oLatencyTrackerDictionary.Count-1
			Dim aLatencyTrackerServers()
			If instr(a(i),"TOTAL") > 0 Then
				iLatencyTrackerServerCount = ubound(aLatencyTrackerServers) + 1
				Redim Preserve aLatencyTrackerServers(iLatencyTrackerServerCount)				
				aLatencyTrackerServers(iLatencyTrackerServerCount) = Array(a(i),oLatencyTrackerDictionary.Item(a(i)))
				'Wscript.echo vbCRLF & "Adding aLatencyTrackerServers entry for " & a(i)
			End If
		Next
		For i = 0 to oLatencyTrackerDictionary.Count-1
			aLatencyTrackerDictionaryEntry = oLatencyTrackerDictionary.Item(a(i))
			aLatencyTrackerDictionaryKey = split(a(i),",")
			For x = 0 to ubound(aLatencyTrackerServers)		
				aLatencyTrackerServersEntry = aLatencyTrackerServers(x)
				If instr(1,aLatencyTrackerServersEntry(0),aLatencyTrackerDictionaryKey(0)) Then
					aLatencyTrackerDictionaryTotalEntry = aLatencyTrackerServersEntry(1)
					'Wscript.echo "Found total (" & aLatencyTrackerDictionaryTotalEntry(0) & ") for " & a(i)
				Else
					'Wscript.echo "Not Found total for " & a(i)
				End If
			Next
			iAvgComponentLatency = round(aLatencyTrackerDictionaryEntry(1)/aLatencyTrackerDictionaryTotalEntry(0),1)
			If instr(1,a(i),"TOTAL") Then
				iComponentPercentileLt1Sec = ""
				iIndividualComponentSLAPercentile = round(100*(aLatencyTrackerDictionaryEntry(2))/aLatencyTrackerDictionaryTotalEntry(0),2)
				iStageComponentSLAPercentile = round(100*(aLatencyTrackerDictionaryEntry(3))/aLatencyTrackerDictionaryTotalEntry(0),2)
				iServerSLAPercentile = round(100*(aLatencyTrackerDictionaryEntry(4))/aLatencyTrackerDictionaryTotalEntry(0),2)
			Else
				iComponentPercentileLt1Sec = round(100*(aLatencyTrackerDictionaryTotalEntry(0)-aLatencyTrackerDictionaryEntry(0))/aLatencyTrackerDictionaryTotalEntry(0),2)
				iIndividualComponentSLAPercentile = round(100*(aLatencyTrackerDictionaryTotalEntry(0)-aLatencyTrackerDictionaryEntry(0)+aLatencyTrackerDictionaryEntry(2))/aLatencyTrackerDictionaryTotalEntry(0),2)
				iStageComponentSLAPercentile = round(100*(aLatencyTrackerDictionaryTotalEntry(0)-aLatencyTrackerDictionaryEntry(0)+aLatencyTrackerDictionaryEntry(3))/aLatencyTrackerDictionaryTotalEntry(0),2)
				iServerSLAPercentile = round(100*(aLatencyTrackerDictionaryTotalEntry(0)-aLatencyTrackerDictionaryEntry(0)+aLatencyTrackerDictionaryEntry(4))/aLatencyTrackerDictionaryTotalEntry(0),2)
			End If
			If a(i) <> "," Then
				oLatencyTrackerResults.WriteLine a(i) & "," & iAvgComponentLatency & "," & iComponentPercentileLt1Sec & "%," & iIndividualComponentSLAPercentile & "%," & iStageComponentSLAPercentile & "%," & iServerSLAPercentile & "%," & aLatencyTrackerDictionaryEntry(5) & "," & aLatencyTrackerDictionaryEntry(6) & "," & aLatencyTrackerDictionaryEntry(7)
			End If
		Next
	End If
	
	iCountUnderCurrentLatency = 0
	oSummaryResults.WriteLine vbcrlf & "Server Latency Distribution for " & iAllServerLatencyTotalCount & " Individual Deliveries:" 
	If iAllServerLatencyTotalCount > 0 Then
		iCountUnderCurrentLatency = iCountUnderCurrentLatency + int(aMsgServerLatency(0,1))
		oSummaryResults.WriteLine "Count of messages 0 Sec =< MsgLatency <= " & aMsgServerLatency(0,0) & " Sec: "  & aMsgServerLatency(0,1) & " (" & round(100*aMsgServerLatency(0,1)/iAllServerLatencyTotalCount,2) & "%, " & round(100*iCountUnderCurrentLatency/iAllServerLatencyTotalCount,2) & "%)"
		For i = 1 to 24
			iCountUnderCurrentLatency = iCountUnderCurrentLatency + int(aMsgServerLatency(i,1))
			oSummaryResults.WriteLine "Count of messages " & aMsgServerLatency(i-1,0) & " Sec < MsgLatency <= " & aMsgServerLatency(i,0) & " Sec: "  & aMsgServerLatency(i,1) & " (" & round(100*aMsgServerLatency(i,1)/iAllServerLatencyTotalCount,2) & "%, " & round(100*iCountUnderCurrentLatency/iAllServerLatencyTotalCount,2) & "%)"
		Next
	End If

	iCountUnderCurrentLatency = 0
	oSummaryResults.WriteLine vbcrlf & "Server Latency Distribution for " & iAllLastDeliveryCount & " Messages Processed:" 
	If iAllLastDeliveryCount > 0 Then
		iCountUnderCurrentLatency = iCountUnderCurrentLatency + int(aMsgMaxServerLatency(0,1))
		oSummaryResults.WriteLine "Count of messages 0 Sec =< MaxMsgLatency <= " & aMsgMaxServerLatency(0,0) & " Sec: "  & aMsgMaxServerLatency(0,1) & " (" & round(100*aMsgMaxServerLatency(0,1)/iAllLastDeliveryCount,2) & "%, " & round(100*iCountUnderCurrentLatency/iAllLastDeliveryCount,2) & "%)"
		For i = 1 to 24
			iCountUnderCurrentLatency = iCountUnderCurrentLatency + int(aMsgMaxServerLatency(i,1))
			oSummaryResults.WriteLine "Count of messages " & aMsgMaxServerLatency(i-1,0) & " Sec < MaxMsgLatency <= " & aMsgMaxServerLatency(i,0) & " Sec: "  & aMsgMaxServerLatency(i,1) & " (" & round(100*aMsgMaxServerLatency(i,1)/iAllLastDeliveryCount,2) & "%, " & round(100*iCountUnderCurrentLatency/iAllLastDeliveryCount,2) & "%)"
		Next
	End If

	
	iCountUnderCurrentLatency = 0
	iLatencyDeliverCount = iDeliverCount - iDeliveryLatencyCountUnknown - iDeliveryLatencyCountLtZero
	oSummaryResults.WriteLine vbcrlf & "End-To-End Delivery Latency Distribution for " & iLatencyDeliverCount & " Messages Delivered:" 
	iCountUnderCurrentLatency = iCountUnderCurrentLatency + int(aMsgDeliveryLatency(0,1))
	If iLatencyDeliverCount > 0 Then
		oSummaryResults.WriteLine "Count of messages 0 Sec =< MsgLatency <= " & aMsgDeliveryLatency(0,0) & " Sec: "  & aMsgDeliveryLatency(0,1) & " (" & round(100*aMsgDeliveryLatency(0,1)/iLatencyDeliverCount,2) & "%, " & round(100*iCountUnderCurrentLatency/iLatencyDeliverCount,2) & "%)"
		For i = 1 to 24
			iCountUnderCurrentLatency = iCountUnderCurrentLatency + int(aMsgDeliveryLatency(i,1))
			oSummaryResults.WriteLine "Count of messages " & aMsgDeliveryLatency(i-1,0) & " Sec < MsgLatency <= " & aMsgDeliveryLatency(i,0) & " Sec: "  & aMsgDeliveryLatency(i,1) & " (" & round(100*aMsgDeliveryLatency(i,1)/iLatencyDeliverCount,2) & "%, " & round(100*iCountUnderCurrentLatency/iLatencyDeliverCount,2) & "%)"
		Next
		iCountUnderCurrentLatency = iCountUnderCurrentLatency + iDeliveryLatencyCountLtZero
		oSummaryResults.WriteLine vbCRLF & "Count of messages with negative latency: " & iDeliveryLatencyCountLtZero & " (" & round(100*iDeliveryLatencyCountLtZero/iDeliverCount,2) & "% of " & iDeliverCount & " total messages delivered)"
		iCountUnderCurrentLatency = iCountUnderCurrentLatency + iDeliveryLatencyCountUnknown
		oSummaryResults.WriteLine "Count of messages with unknown latency: " & iDeliveryLatencyCountUnknown & " (" & round(100*iDeliveryLatencyCountUnknown/iDeliverCount,2) & "% of " & iDeliverCount & " total messages delivered)"
	End If
		
	If iReceiveCount > 0 Then
		oSummaryResults.WriteLine vbcrlf & "Total Messages received: " & iReceiveCount & " (Avg Msg Size: " & round((iTotalMsgBytes/iReceiveCount)/1024,0) & " KB)"
		oSummaryResults.WriteLine "Message Size Distribution can be found in " & sMessageSizeDistribution
		oMessageSizeDistribution.writeline "SizeRange,Count,PercentofTotal,PercentileUnderCurrentSize" 
		iCountUnderCurrentSize = int(aMsgSize(0,1))
		oMessageSizeDistribution.WriteLine "0-" & aMsgSize(0,0)/1024 & " KB"  & "," & aMsgSize(0,1) & "," & round(100*aMsgSize(0,1)/iReceiveCount,2) & "%," & round(100*iCountUnderCurrentSize/iReceiveCount,2) & "%"
		For i = 1 to 16
			iCountUnderCurrentSize = iCountUnderCurrentSize + int(aMsgSize(i,1))
			oMessageSizeDistribution.WriteLine aMsgSize(i-1,0)/1024 & "-" & aMsgSize(i,0)/1024 & " KB" & "," & aMsgSize(i,1) & "," & round(100*aMsgSize(i,1)/iReceiveCount,2) & "%," & round(100*iCountUnderCurrentSize/iReceiveCount,2) & "%"
		Next
		oSummaryResults.WriteLine "Count of messages that exceeds " & iMaxMessageSizeThresholdKB & " KB = " & iExceedsMaxMessageSize
		If iExceedsMaxMessageSize > 0 Then
			oSummaryResults.WriteLine "Details of messages that exceed " & iMaxMessageSizeThresholdKB & " KB at " & sMessageSizeExceptionResults
		Else
			oMessageSizeExceptionResults.close
			If fso.FileExists(sMessageSizeExceptionResults) Then fso.DeleteFile(sMessageSizeExceptionResults)
		End If
	End If
	
	'oDuplicateDeliveryDictionary
	If iDuplicateDeliverRecipCount > 0 Then
		iAvgDuplicateDeliveryRecipients = round(iDuplicateDeliverRecipCount/iDuplicateDeliver,1)
		oSummaryResults.WriteLine vbcrlf & "Total Unique Duplicate Deliveries: " & iDuplicateDeliver
		oSummaryResults.WriteLine "Average Recipients/Duplicate: " & iAvgDuplicateDeliveryRecipients
		oSummaryResults.WriteLine "Duplicate Delivery Details at " & sDuplicateDeliveryResults
		a = oDuplicateDeliveryDictionary.Keys
		oDuplicateDeliveryResults.WriteLine "Sender,MessageId,Recipient,Count"
		For i = 0 to oDuplicateDeliveryDictionary.Count-1		
			If oDuplicateDeliveryDictionary.Item(a(i)) > iAvgDuplicateDeliveryRecipients Then
				oDuplicateDeliveryResults.WriteLine a(i) & "," & oDuplicateDeliveryDictionary.Item(a(i))
			End If
		Next
	End If
	
	'oSubmitDictionary 
	If iSubmitEventCount > 0 Then
		oSummaryResults.WriteLine vbcrlf & "Total MailSubmission Submit Events: " & iSubmitEventCount
		oSummaryResults.WriteLine "MailSubmission ClientType and MessageClass Details at " & sMailSubmissionDistribution
		Set oMailSubmissionDistribution = fso.OpenTextFile(sMailSubmissionDistribution,ForAppending,1)
		a = oSubmitDictionary.Keys
		oMailSubmissionDistribution.WriteLine "ClientHostName,ServerHostName,ClientType,MessageClass,Count"
		For i = 0 to oSubmitDictionary.Count-1		
			oMailSubmissionDistribution.WriteLine a(i) & "," & oSubmitDictionary.Item(a(i))
		Next
	End If
	
	
	'oFinalDeliveryDictionary
	'aFinalDeliveryDictionaryEntry(13)	'0=MsgCount,1=RemainingRecipCount,2=DeliverCount,3=EventId,4=LatencyLtSLA,5=LastEventTime,6=ReturnPath,7=bDumpster,8=TransferCount,9=LastTransferTime,10=iTotalBytes,11=iRecipCountMax,12=ExpandCount,13=DuplicateDeliverCount
	If oFinalDeliveryDictionary.Count > 0 Then
		oSummaryResults.WriteLine vbcrlf & "Final Delivery statistics available in " & sFinalDeliveryResults & vbcrlf
		a = oFinalDeliveryDictionary.Keys
		oFinalDeliveryResults.WriteLine "Source:Event,MsgCount,TotalRecipients,RemainingRecipCount,DeliverCount,AvgDeliveries,%LatencyLt30,%Dumpster,TransferCount,AvgBytesPerMsg,AvgRecipCount,ExpandCount,%DuplicateDeliver"
		'oFinalDeliveryResults.WriteLine "Source:Event,0=MsgCount,1=RemainingRecipCount,2=DeliverCount,3=EventId,4=LatencyLtSLA,5=LastEventTime,6=ReturnPath,7=bDumpster,8=TransferCount,9=LastTransferTime,10=iTotalBytes,11=iRecipCountMax,12=ExpandCount,13=DuplicateDeliverCount"
		For i = 0 to oFinalDeliveryDictionary.Count-1		
			aFinalDeliveryDictionaryEntry = oFinalDeliveryDictionary.Item(a(i))
			sFinalDeliveryDictionaryEntry = NULL
			'For x = 0 to 13
			'	sFinalDeliveryDictionaryEntry = sFinalDeliveryDictionaryEntry & aFinalDeliveryDictionaryEntry(x) & ","
			'Next
			'oFinalDeliveryResults.WriteLine a(i) & "," & sFinalDeliveryDictionaryEntry
			oFinalDeliveryResults.WriteLine a(i) & "," & aFinalDeliveryDictionaryEntry(0) & "," & aFinalDeliveryDictionaryEntry(11) & "," & aFinalDeliveryDictionaryEntry(1) & "," & aFinalDeliveryDictionaryEntry(2) & "," & round(aFinalDeliveryDictionaryEntry(2)/aFinalDeliveryDictionaryEntry(0),1) & "," & round(100*aFinalDeliveryDictionaryEntry(4)/aFinalDeliveryDictionaryEntry(0),1) & "%," & round(100*aFinalDeliveryDictionaryEntry(7)/aFinalDeliveryDictionaryEntry(0),0) & "%," & aFinalDeliveryDictionaryEntry(8) & "," & round(aFinalDeliveryDictionaryEntry(10)/aFinalDeliveryDictionaryEntry(0),0) & "," & round(aFinalDeliveryDictionaryEntry(11)/aFinalDeliveryDictionaryEntry(0),1) & "," & aFinalDeliveryDictionaryEntry(12) & "," & round(100*aFinalDeliveryDictionaryEntry(13)/aFinalDeliveryDictionaryEntry(2),0) & "%"
			
		Next
	End If
	

	LogText vbCRLF & "Summary Results available in " & sSummaryResults
		
End Sub

Function GetProcessInfo(byVal bIncludeCmdLine)

	On Error Resume Next
	strNameOfUser = ""
	Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" _
			& "." & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery( _
			"Select * from Win32_Process " _
			& "Where Name = 'cscript.exe'" _
			& "OR Name = 'wscript.exe'",,48) 
			
	sRuntime = datediff("s",sStartedScript,now)

	For Each objProcess in colProcesses
		If instr(1,lcase(objProcess.CommandLine),"processtrackinglog.vbs") Then		
			Return = objProcess.GetOwner(strNameOfUser)
			sngProcessTime = (CSng(objProcess.KernelModeTime) + CSng(objProcess.UserModeTime)) / 10000000
			If bIncludeCmdLine Then
				GetProcessInfo = vbcrlf & "Process Summary: """ & objProcess.CommandLine & """ (PID=" & objProcess.ProcessId & ")" & vbtab & "WorkingSet=" & cdbl(objProcess.WorkingSetSize) & vbtab & "RunTime=" & sRuntime & vbtab & "ProcessTime=" & sngProcessTime & vbtab & "Files=" & iEnterParseFileFunctionTotal & vbtab & "UserName=""" & strNameOfUser & """"
			Else
				GetProcessInfo = vbcrlf & "Process Summary: " & objProcess.Name & "(PID=" & objProcess.ProcessId & ")" & vbtab & "WorkingSet=" & cdbl(objProcess.WorkingSetSize) & vbtab & "RunTime=" & sRuntime & vbtab & "ProcessTime=" & sngProcessTime & vbtab & "Files=" & iEnterParseFileFunctionTotal & vbtab & "UserName=""" & strNameOfUser & """"
			End If
		End If
	Next

End Function

Sub LogText(byVal sText)
	wscript.echo sText
	oRunTimeLog.writeline sText
	If instr(1,sText,"ERROR") OR instr(1,sText,"WARNING") Then
	 oLogError.writeline sText
	End If
End Sub

Sub ArchiveFiles
	On Error Resume Next
	sYear = year(date)
	sMonth = month(date)
	If len(sMonth) < 2 Then sMonth = "0" & sMonth
	sDay = day(date)
	If len(sDay) < 2 Then sDay = "0" & sDay
	sHour = hour(now)
	If len(sHour) < 2 Then sHour = "0" & sHour
	sMin = minute(now)
	If len(sMin) < 2 Then sMin = "0" & sMin
	sSec = second(now)
	If len(sSec) < 2 Then sSec = "0" & sSec
	sTimestamp = sYear & sMonth & sDay & sHour & sMin & sSec
	If strFilterDate = "" Then
		sArchiveFilename = sArchivePath & "ProcessTrackingLogResults_" & sTimestamp & "_" & strRole & ".zip"
	Else
		sArchiveFilename = sArchivePath & "ProcessTrackingLogResults_" & sTimestamp & "_" & strRole & "_" & replace(strFilterDate,"/","") & ".zip"
	End If
	If fso.FileExists(sArchiveFilename) Then 
		LogText "ERROR: could not archive files using file " & sArchiveFilename
		Exit Sub
	End If
	strExePath = strProgramFilesPath & "\WinZip\WZZIP.EXE"
	If fso.FileExists(strExePath) Then
		strCommand = strExePath & " " & sArchiveFilename & " " & sResultPath & "*.*"
		LogText vbcrlf & "Calling """ & strCommand & """"
		strStartArchive = now
		Set oExec = WshShell.Exec(strCommand)

		Do While oExec.Status = 0
			WScript.Sleep 1000
			iArchiveRunTimeSec = DateDiff("s",strStartArchive,now)
			If iArchiveRunTimeSec Mod 30 = 0 Then 
				If DateDiff("s",strStartArchive,now)>=600 Then 
					LogText "ERROR: Timeout occured while attempting to archive files"				
					Exit Sub
				Else
					LogText "WARNING: attempt to archive files has taken " & iArchiveRunTimeSec & " sec (will timeout at 600 sec)"								
				End If
			End if			

		Loop		

		If Not oExec.StdOut.AtEndOfStream Then
			LogText vbCRLF & oExec.StdOut.ReadAll
		End If	
		
	Else
		LogText "Could not find " & strExePath
	End If
		
End Sub

Sub CloseFiles

	oSummaryResults.close
	'oRunTimeLog.close
	'oLogError.close
	oNextHopResults.close
	oReceiveClientFileOut.close
	oLogStatistics.close
	oDsnFailureResults.close
	If bWriteDsnFailureNotFoundResults Then oDsnFailureNotFoundResults.close
	oMbxFullRecipResults.close	
	oTopSendersbySubmitResults.close
	oTopSendersbyDeliverResults.close
	oTopRecipientResults.close
	oExpandFileOut.close
	oDeliveryLatencyExceptionResults.close
	oMessageSizeExceptionResults.close
	oFinalDeliveryResults.close

End Sub

Function UpdateServerLatency(byRef aFileLine, byRef bServerLatencySLAMet)
	On Error Resume Next
	UpdateServerLatency = FALSE
	'Exit Function
	'bDebugOut = TRUE
	strUniqueMessageid = aFileLine(iInternalMsgId) & "," & aFileLine(iMessageId)
	sCurrentEventTime=GetDateTime(aFileLine(iEventTime))
	If oServerLatencyDictionary.Exists(strUniqueMessageId) Then
		aLatencyTrackingEntry = oServerLatencyDictionary.Item(strUniqueMessageId)			
		iServerLatency = datediff("s",aLatencyTrackingEntry(0),sCurrentEventTime)
		aLatencyTrackingEntry(4) = iServerLatency		
		aLatencyTrackingEntry(5) = sCurrentEventTime

		If instr(1,"DELIVER,SEND,FAIL,POISONMESSAGE,DUPLICATEDELIVER",ucase(aFileLine(iEventId))) Then
			
			iServerLatencyTotalCount = iServerLatencyTotalCount + 1
			
			If ucase(aFileLine(iEventId)) = "DELIVER" AND NOT(aLatencyTrackingEntry(7)) Then
				'Wscript.echo aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5) & vbtab & aLatencyTrackingEntry(6) & vbtab & aLatencyTrackingEntry(7)
				aLatencyTrackingEntry(7) = TRUE
				oServerLatencyDictionary.Item(strUniqueMessageId) = aLatencyTrackingEntry
			End If		
			
			If ucase(aFileLine(iEventId)) = "DUPLICATEDELIVER" Then
				aLatencyTrackingEntry(13) = aLatencyTrackingEntry(13) + 1
			End If
			
			iServerLatencyRecipientsCounted = iServerLatencyRecipientsCounted + int(aFileLine(iRecipCount))
						
			If 0 =< iServerLatency and iServerLatency <= 1 Then
					aMsgServerLatency(0,1) = int(aMsgServerLatency(0,1)) + 1
			Else
				For i = 1 to 24 
					If 2^(i-1) < iServerLatency and iServerLatency <= (2^i) Then
						aMsgServerLatency(i,1) = int(aMsgServerLatency(i,1)) + 1
					End If
				Next				                
			End If
			
			If iServerLatency <= iServerLatencySLA Then
				bServerLatencySLAMet = TRUE
				iServerLatencySLAMetCount = iServerLatencySLAMetCount + 1
			End If
									
			'If bDebugOut Then Wscript.echo "Evaluating oServerLatencyDictionary.Item = " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)

			If ucase(aFileLine(iEventId)) = "BADMAIL" Then
				'iServerLatencyRecipientsCounted = iServerLatencyRecipientsCounted + int(aFileLine(iRecipCount))
				'If bDebugOut Then Wscript.echo "Removing oServerLatencyDictionary.Item = " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)
				'TrackLastDelivery aLatencyTrackingEntry, aFileLine(iEventId), ""
				'oServerLatencyDictionary.Remove strUniqueMessageId
			Else
				aLatencyTrackingEntry(1) = aLatencyTrackingEntry(1) - int(aFileLine(iRecipCount))
				aLatencyTrackingEntry(2) = aLatencyTrackingEntry(2) + 1
				If aLatencyTrackingEntry(1) <= 0 Then
					'If bDebugOut Then Wscript.echo "Removing oServerLatencyDictionary.Item = " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)
					TrackLastDelivery aLatencyTrackingEntry, aFileLine(iEventId), aFileLine(iMessageInfo)
					oServerLatencyDictionary.Remove strUniqueMessageId
				Else
					'If bDebugOut Then Wscript.echo "Updating oServerLatencyDictionary.Item = " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)
					oServerLatencyDictionary.Item(strUniqueMessageId) = aLatencyTrackingEntry
				End If
			End If
			
			'Parse MessageInfo field if containing E14 Latency info
			'2007-10-24T00:11:42.660Z;SRV=FQDN1:TOTAL=<Sec>|Component=<sec>|Component2=<sec>;SRV=FQDN2:TOTAL=<Sec>|Component=<sec>|Component2=<sec>
			If instr(1,aFileLine(iMessageInfo),";") Then
				aMessageInfo = split(aFileLine(iMessageInfo),";") 
				'Only look at last latency tracking information
				iLastTracker = ubound(aMessageInfo)			
				'Wscript.echo iLastTracker & vbtab & aMessageInfo(iLastTracker)
				aLatencyTrackerEntry = split(aMessageInfo(iLastTracker),":")		
				strLatencyTrackerFqdn = right(aLatencyTrackerEntry(0),len(aLatencyTrackerEntry(0)) - instr(aLatencyTrackerEntry(0),"="))
				aLatencyTrackerComponentEntry = split(aLatencyTrackerEntry(1),"|")
				For x = 0 to ubound(aLatencyTrackerComponentEntry)					
					aLatencyTrackerComponentEntryDetails = split(aLatencyTrackerComponentEntry(x),"=")
					
					sLatencyTrackerComponentDictionaryKey = aLatencyTrackerComponentEntryDetails(0)
					If oLatencyTrackerComponentDictionary.Exists(sLatencyTrackerComponentDictionaryKey) Then
						bNewComponent = FALSE
						aLatencyTrackerComponentDictionaryEntry2 = oLatencyTrackerComponentDictionary.Item(sLatencyTrackerComponentDictionaryKey)
						aLatencyTrackerComponentDictionaryEntry(0) = aLatencyTrackerComponentDictionaryEntry2(0) + 1
						aLatencyTrackerComponentDictionaryEntry(1) = aLatencyTrackerComponentDictionaryEntry2(2)
						aLatencyTrackerComponentDictionaryEntry(2) = aLatencyTrackerComponentDictionaryEntry2(2) + int(aLatencyTrackerComponentEntryDetails(1))
						oLatencyTrackerComponentDictionary.Item(sLatencyTrackerComponentDictionaryKey) = aLatencyTrackerComponentDictionaryEntry
						'If err.number <> 0 Then Wscript.echo "ERROR, " & err.Description
					Else
						bNewComponent = TRUE
						ReDim aLatencyTrackerComponentDictionaryEntry(2) '0=ComponentUsageCount,1=PreviousLatency,2=TotalLatency
						aLatencyTrackerComponentDictionaryEntry(0) = 1
						aLatencyTrackerComponentDictionaryEntry(1) = 0
						aLatencyTrackerComponentDictionaryEntry(2) = int(aLatencyTrackerComponentEntryDetails(1))
						oLatencyTrackerComponentDictionary.Add sLatencyTrackerComponentDictionaryKey,aLatencyTrackerComponentDictionaryEntry
					End If
					
					
					If bNewComponent Then
						'Wscript.echo x & "," & oLatencyTrackerComponentDictionary.Count & "," & sLatencyTrackerComponentDictionaryKey & "," & aLatencyTrackerComponentDictionaryEntry(0) & "," & aLatencyTrackerComponentDictionaryEntry(1) & "," & aLatencyTrackerComponentDictionaryEntry(2) & "," & cstr(bNewComponent)
					Else
						'Wscript.echo x & "," & oLatencyTrackerComponentDictionary.Count & "," & sLatencyTrackerComponentDictionaryKey & "," & aLatencyTrackerComponentDictionaryEntry(0) & "," & aLatencyTrackerComponentDictionaryEntry(1) & "," & aLatencyTrackerComponentDictionaryEntry(2) & "," & cstr(bNewComponent)
					End If
					
					sLatencyTrackerDictionaryKey = strLatencyTrackerFqdn & "," & aLatencyTrackerComponentEntryDetails(0)
					If strLatencyTrackerFqdn <> "" AND aLatencyTrackerComponentEntryDetails(0) <> "" Then
						If oLatencyTrackerDictionary.Exists(sLatencyTrackerDictionaryKey) Then
							aLatencyTrackerDictionaryEntry = oLatencyTrackerDictionary.Item(sLatencyTrackerDictionaryKey)
							If bNewComponent Then aLatencyTrackerDictionaryEntry(0) = aLatencyTrackerDictionaryEntry(0) + 1
							aLatencyTrackerDictionaryEntry(1) = aLatencyTrackerDictionaryEntry(1) + int(aLatencyTrackerComponentEntryDetails(1))
							'Create Individual Component Latency SLA Percentile
							If aLatencyTrackerComponentDictionaryEntry(2) <= iIndividualComponentLatencySLA Then
								aLatencyTrackerDictionaryEntry(2) = aLatencyTrackerDictionaryEntry(2) + 1
							End If
							'Create Stage Component Latency SLA Percentile
							If aLatencyTrackerComponentDictionaryEntry(2) <= iStageComponentLatencySLA Then
								aLatencyTrackerDictionaryEntry(3) = aLatencyTrackerDictionaryEntry(3) + 1
							End If							
							'Create Server Latency SLA Percentile
							If aLatencyTrackerComponentDictionaryEntry(2) <= iServerLatencySLA Then
								aLatencyTrackerDictionaryEntry(4) = aLatencyTrackerDictionaryEntry(4) + 1
							End If
							If Not(bNewComponent) Then
								'Update Latency SLA Percentiles (decrement count based on previous latency total)
								If aLatencyTrackerComponentDictionaryEntry(1) <= iIndividualComponentLatencySLA Then
									aLatencyTrackerDictionaryEntry(2) = aLatencyTrackerDictionaryEntry(2) - 1
								End If
								If aLatencyTrackerComponentDictionaryEntry(1) <= iStageComponentLatencySLA Then
									aLatencyTrackerDictionaryEntry(3) = aLatencyTrackerDictionaryEntry(3) - 1
								End If							
								If aLatencyTrackerComponentDictionaryEntry(1) <= iServerLatencySLA Then
									aLatencyTrackerDictionaryEntry(4) = aLatencyTrackerDictionaryEntry(4) - 1
								End If							
							End If
							aLatencyTrackerDictionaryEntry(5) = aLatencyTrackerDictionaryEntry(5) + 1
							If aLatencyTrackerComponentDictionaryEntry(0) > aLatencyTrackerDictionaryEntry(6) Then
								aLatencyTrackerDictionaryEntry(6) = aLatencyTrackerComponentDictionaryEntry(0)
							End If
							If aLatencyTrackerComponentDictionaryEntry(2) > aLatencyTrackerDictionaryEntry(7) Then
								aLatencyTrackerDictionaryEntry(7) = aLatencyTrackerComponentDictionaryEntry(2)
							End If							
							oLatencyTrackerDictionary.Item(sLatencyTrackerDictionaryKey) = aLatencyTrackerDictionaryEntry
						Else
							Dim aLatencyTrackerDictionaryEntry(7) '0=MessageCountGt1Sec,1=TotalLatency,2=IndividualComponentSLA,3=StageComponentSLA,4=ServerSLA,5=TotalInvocationsGt1Sec,6=MaxInvocationCount,7=MaxLatency
							aLatencyTrackerDictionaryEntry(0) = 1
							aLatencyTrackerDictionaryEntry(1) = int(aLatencyTrackerComponentEntryDetails(1))
							'Create Individual Component Latency SLA Percentile
							If int(aLatencyTrackerComponentEntryDetails(1)) <= iIndividualComponentLatencySLA Then
								aLatencyTrackerDictionaryEntry(2) = 1
							Else
								aLatencyTrackerDictionaryEntry(2) = 0
							End If
							'Create Stage Component Latency SLA Percentile
							If int(aLatencyTrackerComponentEntryDetails(1)) <= iStageComponentLatencySLA Then
								aLatencyTrackerDictionaryEntry(3) = 1
							Else
								aLatencyTrackerDictionaryEntry(3) = 0
							End If							
							'Create Server Latency SLA Percentile
							If int(aLatencyTrackerComponentEntryDetails(1)) <= iServerLatencySLA Then
								aLatencyTrackerDictionaryEntry(4) = 1
							Else
								aLatencyTrackerDictionaryEntry(4) = 0
							End If
							aLatencyTrackerDictionaryEntry(5) = 1
							aLatencyTrackerDictionaryEntry(6) = 1
							aLatencyTrackerDictionaryEntry(7) = int(aLatencyTrackerComponentEntryDetails(1))
							oLatencyTrackerDictionary.Add sLatencyTrackerDictionaryKey,aLatencyTrackerDictionaryEntry
						End If
					End If
				Next
				oLatencyTrackerComponentDictionary.RemoveAll
			End If

									
		ElseIf instr(1,"EXPAND",ucase(aFileLine(iEventId))) Then
			iExpandRecipients = int(aFileLine(iRecipCount))			
			If instr(1, aFileline(iRelatedRecipAddr), ";") Then
				aRelatedRecipients = split(aFileline(iRelatedRecipAddr),";")
				iRelatedRecipients = ubound(aRelatedRecipients) + 1
			Else
				iRelatedRecipients = 	1
			End If
			aLatencyTrackingEntry(1) = aLatencyTrackingEntry(1) + iExpandRecipients - 1
			aLatencyTrackingEntry(11) = aLatencyTrackingEntry(11) + iExpandRecipients - 1
			aLatencyTrackingEntry(12) = aLatencyTrackingEntry(12) + 1
			If bDebugOut Then Wscript.echo "Updating oServerLatencyDictionary.Item = " & strUniqueMessageId & vbtab & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)
			If aLatencyTrackingEntry(1) <= 0 Then
				If bDebugOut Then Wscript.echo "Removing oServerLatencyDictionary.Item = " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3) & vbtab & aLatencyTrackingEntry(4) & vbtab & aLatencyTrackingEntry(5)
				TrackLastDelivery aLatencyTrackingEntry, aFileLine(iEventId), ""
				oServerLatencyDictionary.Remove strUniqueMessageId
			Else
				oServerLatencyDictionary.Item(strUniqueMessageId) = aLatencyTrackingEntry
			End If
		
		ElseIf instr(1,"BADMAIL",ucase(aFileLine(iEventId))) Then	
			If bDebugOut Then Wscript.echo "Found in queue, ignorning " & aFileLine(iEventId) & "," & strUniqueMessageId
		
		ElseIf instr(1,"RESOLVE,REDIRECT",ucase(aFileLine(iEventId))) Then
			If bDebugOut Then Wscript.echo "Found in queue, ignorning " & aFileLine(iEventId) & "," & strUniqueMessageId
		
		Else			
			If bDebugOut Then Wscript.echo "Found in queue, ignorning " & aFileLine(iEventId) & "," & strUniqueMessageId
											
		End If
					
	ElseIf instr(1,"RECEIVE,DSN,TRANSFER,RESUBMIT",ucase(aFileLine(iEventId)))>0 Then
		Dim aLatencyTrackingEntry(13)	'0=ReceiveTime,1=RecipCount,2=DeliverCount,3=EventId,4=Latency,5=LastEventTime,6=ReturnPath,7=bDumpster,8=TransferCount,9=LastTransferTime,10=iTotalBytes,11=iRecipCountMax,12=ExpandCount,13=DuplicateDeliverCount
		aLatencyTrackingEntry(0) = sCurrentEventTime											
		aLatencyTrackingEntry(1) = int(aFileLine(iRecipCount))
		aLatencyTrackingEntry(2) = 0
		aLatencyTrackingEntry(3) = ucase(aFileLine(iSource)) & ":" & ucase(aFileLine(iEventId))
		aLatencyTrackingEntry(4) = 0
		aLatencyTrackingEntry(5) = NULL
		aLatencyTrackingEntry(6) = aFileLine(iReturnPath)
		aLatencyTrackingEntry(7) = FALSE
		aLatencyTrackingEntry(8) = 0
		aLatencyTrackingEntry(9) = NULL
		aLatencyTrackingEntry(10) = int(aFileLine(iTotalBytes))
		aLatencyTrackingEntry(11) = int(aFileLine(iRecipCount))
		aLatencyTrackingEntry(12) = 0
		aLatencyTrackingEntry(13) = 0
		
		If aLatencyTrackingEntry(3) = "AGENT:RECEIVE" Then
			aLatencyTrackingEntry(3) = aLatencyTrackingEntry(3) & "(" & aFileLine(iSourceContext) & ")"
		End If
		
		If aLatencyTrackingEntry(1) <> 0 Then
		
			Select Case aLatencyTrackingEntry(6) 
			
				Case "<>"

					Select Case ucase(aFileLine(iSource))
						
						Case "STOREDRIVER"
						
							If instr(1,aFileLine(iSubject),"Meeting Forward Notification:")>0 or instr(1,aFileLine(iSubject),"Message Recall")>0 Then
								CalculateQueueSize sCurrentEventTime, True, aLatencyTrackingEntry(10)
								oServerLatencyDictionary.Add strUniqueMessageId,aLatencyTrackingEntry									
							
							ElseIf instr(1,aFileLine(iSubject),"Out of Office:") > 0 Then
								'Don't add OOF messages because of 'AckStatus.SuccessNoDsn' issue with OOF suppression
								'Wscript.echo " >>Won't add oServerLatencyDictionary entry (OOF)"
							
							ElseIf instr(1,aFileLine(iSubject),":") > 0 Then
								'Don't add OOF messages because of 'AckStatus.SuccessNoDsn' issue (foreign OOF can't be detected)
								'LogText " >>Won't add oServerLatencyDictionary entry " & aFileline(iSubject)
							
							Else
								CalculateQueueSize sCurrentEventTime, True, aLatencyTrackingEntry(10)
								oServerLatencyDictionary.Add strUniqueMessageId,aLatencyTrackingEntry									
							
							End If
							
						Case Else
							CalculateQueueSize sCurrentEventTime, True, aLatencyTrackingEntry(10)
							oServerLatencyDictionary.Add strUniqueMessageId,aLatencyTrackingEntry			
							
					End Select
				
				Case Else
					CalculateQueueSize sCurrentEventTime, True, aLatencyTrackingEntry(10)
					oServerLatencyDictionary.Add strUniqueMessageId,aLatencyTrackingEntry

				
			End Select

		Else
			LogText "Attempting to add invalid oServerLatencyDictionary entry " & aLatencyTrackingEntry(0) & vbtab & aLatencyTrackingEntry(1) & vbtab & aLatencyTrackingEntry(2) & vbtab & aLatencyTrackingEntry(3)
			LogText ">>>" & aFileLine(iEventId) & "," & aFileLine(iSource)
			Exit Function
			
		End If
		
		If instr(1,"TRANSFER",ucase(aFileLine(iEventId))) Then	
			
			aLatencyTrackingEntry(3) = ucase(aFileLine(iSource)) & ":" & ucase(aFileLine(iEventId)) & "(" & aFileLine(iSourceContext) & ")"
			oServerLatencyDictionary.Item(strUniqueMessageId) = aLatencyTrackingEntry
									
			strAltUniqueMessageId = aFileLine(iReference) & "," & aFileLine(iMessageId)
			'Wscript.echo "Lookup " & (aFileLine(iEventId)) & " TMID " & aFileLine(iInternalMsgId) & " using " & strAltUniqueMessageId
			
			If oServerLatencyDictionary.Exists(strAltUniqueMessageId) Then				
								
				aLatencyTrackingEntry2 = oServerLatencyDictionary.Item(strAltUniqueMessageId)	
				
				iServerLatency = datediff("s",aLatencyTrackingEntry2(0),sCurrentEventTime)
					
				'Wscript.echo "Found " & aLatencyTrackingEntry2(3) & vbtab & strAltUniqueMessageId
				aLatencyTrackingEntry2(1) = aLatencyTrackingEntry2(1) - int(aFileLine(iRecipCount))
				aLatencyTrackingEntry2(4) = iServerLatency		
				aLatencyTrackingEntry2(5) = sCurrentEventTime
				aLatencyTrackingEntry2(8) = aLatencyTrackingEntry2(8) + 1
				aLatencyTrackingEntry2(9) = sCurrentEventTime
				iPreTransferLatency = datediff("s",aLatencyTrackingEntry2(0),aLatencyTrackingEntry(0)) 
				iTotalPreTransferLatency = iTotalPreTransferLatency + iPreTransferLatency				
				
				If iPreTransferLatency > iPreTransferLatencySLA Then
					iPreTransferLatencyGtSLA = iPreTransferLatencyGtSLA + 1
					'"ReceiveTime,TransferTime,MessageId,SourceContext,Latency"
					'oTransferLatencyExceptionResults.writeline aLatencyTrackingEntry2(0) & "," & aLatencyTrackingEntry(0) & "," & aFileLine(iMessageId) & "," & aFileLine(iSourceContext) & "," & iPreTransferLatency
				End If				
											
				If aLatencyTrackingEntry2(1) <= 0 Then
					'If bDebugOut Then Wscript.echo "Removing oServerLatencyDictionary entry for " & strAltUniqueMessageId & vbtab & aLatencyTrackingEntry2(0) & vbtab & aLatencyTrackingEntry2(1) & vbtab & aLatencyTrackingEntry2(2) & vbtab & aLatencyTrackingEntry2(3) & vbtab & aLatencyTrackingEntry2(4) & vbtab & aLatencyTrackingEntry2(5)
					TrackLastDelivery aLatencyTrackingEntry2, aFileLine(iEventId) & "(" & aFileLine(iSourceContext) & ")", ""
					oServerLatencyDictionary.Remove strAltUniqueMessageId
				Else
					'If bDebugOut Then Wscript.echo "Updating oServerLatencyDictionary.Item = " & strAltUniqueMessageId & vbtab & aLatencyTrackingEntry2(0) & vbtab & aLatencyTrackingEntry2(1) & vbtab & aLatencyTrackingEntry2(2) & vbtab & aLatencyTrackingEntry2(3) & vbtab & aLatencyTrackingEntry2(4) & vbtab & aLatencyTrackingEntry2(5)
					oServerLatencyDictionary.Item(strAltUniqueMessageId) = aLatencyTrackingEntry2
				End If					
			
			Else
				
				'If bDebugOut Then Wscript.echo "Could not find TRANSFER(" & aFileLine(iInternalMsgId) & ") reference = " & strAltUniqueMessageId
				
			End If
			
		End If 'oServerLatencyDictionary.Exists false and "TRANSER"
					
	Else 'oServerLatencyDictionary.Exists false and not "RECEIVE,DSN,TRANSFER,RESUBMIT"
		
		If bDebugOut Then wscript.echo "No entry in dictionary, not tracking " & aFileLine(iEventId) & "," & strUniqueMessageId
		
	End If 'oServerLatencyDictionary.Exists check

	UpdateServerLatency = TRUE
	bDebugOut = FALSE
	
End Function

Function TrackNextHopLatency(byRef aFileLine, byVal bServerLatencySLAMet, byVal bDeliveryLatencyMet)
	On Error Resume Next
	TrackNextHopLatency = FALSE
	
	'Removed aFileLine(iConnectorId)<>"Intra-Organization SMTP Send Connector" (now includes this string in strMySource)
	If aFileLine(iSource) = "SMTP" Then
		strMySource = "SMTP(" & replace(aFileLine(iConnectorId),":","_") & ")"
	Else
		strMySource = aFileLine(iSource)
	End If
	
	If aFileLine(iServerName) <> "" Then
		sNextHopServerDictionaryKey = lcase(aFileLine(iClientName)) & ":" & strMySource & ":" & lcase(aFileLine(iServerName))
	Else
		sNextHopServerDictionaryKey = lcase(aFileLine(iClientName)) & ":" & strMySource & ":" & aFileLine(iServerIP)
	End If
		
	If oNextHopServerDictionary.Exists(sNextHopServerDictionaryKey) Then
		
		aCount = oNextHopServerDictionary.Item(sNextHopServerDictionaryKey)
		aCount(0) = aCount(0) + 1
		aCount(1) = aCount(1) + int(aFileLine(iTotalBytes))		
		If bServerLatencySLAMet Then 
			aCount(2) = aCount(2) + 1
		End If
		If bDeliveryLatencyMet Then
			aCount(3) = aCount(3) + 1
		End If
		oNextHopServerDictionary.Item(sNextHopServerDictionaryKey) = aCount
		
	Else
		
		Dim aCount(3)
		aCount(0) = 1
		aCount(1) = int(aFileLine(iTotalBytes))
		If bServerLatencySLAMet Then 
			aCount(2) = 1
		End If
		If bDeliveryLatencyMet Then
			aCount(3) = 1
		Else
			aCount(3) = 0
		End If
		oNextHopServerDictionary.Add sNextHopServerDictionaryKey,aCount

	End If
	TrackNextHopLatency = TRUE
	
End Function

Sub TrackLastDelivery (byRef aFinalLatencyTrackingEntry, byRef sFinalEventId, byRef sMessageInfo)
	CalculateQueueSize aFinalLatencyTrackingEntry(5), false, aFinalLatencyTrackingEntry(10)
	iFileLastDeliveryCount = iFileLastDeliveryCount + 1
	If aFinalLatencyTrackingEntry(7) Then
		iFileDumpsterTrueCount = iFileDumpsterTrueCount + 1	
	End If
	'Keeps track of latency of last delivery (or transfer)
	iAllLastDeliveryCount = iAllLastDeliveryCount + 1
	If 0 =< aFinalLatencyTrackingEntry(4) and aFinalLatencyTrackingEntry(4) <= 1 Then
			aMsgMaxServerLatency(0,1) = int(aMsgMaxServerLatency(0,1)) + 1
	Else
		For i = 1 to 24 
			If 2^(i-1) < aFinalLatencyTrackingEntry(4) and aFinalLatencyTrackingEntry(4) <= (2^i) Then
				aMsgMaxServerLatency(i,1) = int(aMsgMaxServerLatency(i,1)) + 1
			End If
		Next				                
	End If
	
	If instr(1,aFinalLatencyTrackingEntry(3),"TRANSFER")>0 AND instr(1,sFinalEventId,"TRANSFER")>0 Then
		'ReceiveTime,RecipCount,DeliverCount,EventId,Latency,LastEventTime,ReturnPath,bDumpster,TransferCount,LastTransferTime	
		'Wscript.echo "FinalLatency: " & aFinalLatencyTrackingEntry(0) & vbtab & aFinalLatencyTrackingEntry(1) & vbtab & aFinalLatencyTrackingEntry(2)& vbtab & aFinalLatencyTrackingEntry(3)& vbtab & aFinalLatencyTrackingEntry(4) & vbtab & aFinalLatencyTrackingEntry(5) & vbtab & aFinalLatencyTrackingEntry(8) & vbtab & aFinalLatencyTrackingEntry(9) & vbtab & sFinalEventId
	ElseIf sFinalEventId="CLEANUP" Then
		WriteOrphanQueueEntry strUniqueMessageId, aFinalLatencyTrackingEntry
		'Wscript.echo "FinalLatency: " & aFinalLatencyTrackingEntry(0) & vbtab & aFinalLatencyTrackingEntry(1) & vbtab & aFinalLatencyTrackingEntry(2)& vbtab & aFinalLatencyTrackingEntry(3)& vbtab & aFinalLatencyTrackingEntry(4) & vbtab & aFinalLatencyTrackingEntry(5) & vbtab & aFinalLatencyTrackingEntry(7) & vbtab & aFinalLatencyTrackingEntry(8) & vbtab & aFinalLatencyTrackingEntry(9) & vbtab & sFinalEventId
	ElseIf instr(1,aFinalLatencyTrackingEntry(3),"RESUBMIT")>0 Then
		
		strOriginalArrivalTime = GetOriginalArrivalTime(sMessageInfo)
		If strOriginalArrivalTime = "" Then			
			Wscript.echo "Error, strOriginalArrivalTime is blank"
		Else
			strResubmitMesage = sFinalEventId & ": Message received at " & strOriginalArrivalTime & " UTC and resubmitted at " & aFinalLatencyTrackingEntry(0) & " UTC"
			If datediff("s",strOriginalArrivalTime,aFinalLatencyTrackingEntry(0)) > 2*3600 Then			
				LogText now & vbtab & "INFO" & vbtab & strResubmitMesage
			End If
		End If

		Select case sFinalEventId
			Case "DUPLICATEDELIVER"
				'Wscript.echo "RESUBMIT DUPLICATE DELIVER"
			Case Else
				LogText now & vbtab & "INFO" & vbtab & strResubmitMesage
				iResubmitDeliveryCount = iResubmitDeliveryCount + 1
		End Select
	End If
	'If oTransferDictionary.Exists(sTransferDictionaryKey)
	
	If aFinalLatencyTrackingEntry(11) > 0 Then
		sFinalDeliveryDictionaryKey = aFinalLatencyTrackingEntry(3)
		If oFinalDeliveryDictionary.Exists(sFinalDeliveryDictionaryKey) Then
			aFinalDeliveryDictionaryEntry = oFinalDeliveryDictionary.Item(sFinalDeliveryDictionaryKey)
			aFinalDeliveryDictionaryEntry(0) = aFinalDeliveryDictionaryEntry(0) + 1
			aFinalDeliveryDictionaryEntry(1) = aFinalDeliveryDictionaryEntry(1) + aFinalLatencyTrackingEntry(1)
			aFinalDeliveryDictionaryEntry(2) = aFinalDeliveryDictionaryEntry(2) + aFinalLatencyTrackingEntry(2)
			If aFinalLatencyTrackingEntry(4) <= iServerLatencySLA Then aFinalDeliveryDictionaryEntry(4) = aFinalDeliveryDictionaryEntry(4) + 1
			If aFinalLatencyTrackingEntry(7) Then aFinalDeliveryDictionaryEntry(7) = aFinalDeliveryDictionaryEntry(7) + 1			
			aFinalDeliveryDictionaryEntry(8) = aFinalDeliveryDictionaryEntry(8) + aFinalLatencyTrackingEntry(8)
			aFinalDeliveryDictionaryEntry(10) = aFinalDeliveryDictionaryEntry(10) + aFinalLatencyTrackingEntry(10)
			aFinalDeliveryDictionaryEntry(11) = aFinalDeliveryDictionaryEntry(11) + aFinalLatencyTrackingEntry(11)
			aFinalDeliveryDictionaryEntry(12) = aFinalDeliveryDictionaryEntry(12) + aFinalLatencyTrackingEntry(12)
			aFinalDeliveryDictionaryEntry(13) = aFinalDeliveryDictionaryEntry(13) + aFinalLatencyTrackingEntry(13)
			oFinalDeliveryDictionary.Item(sFinalDeliveryDictionaryKey) = aFinalDeliveryDictionaryEntry
		Else
			Dim aFinalDeliveryDictionaryEntry(13)	'0=MsgCount,1=RemainingRecipCount,2=DeliverCount,3=EventId,4=LatencyLtSLA,5=LastEventTime,6=ReturnPath,7=bDumpster,8=TransferCount,9=LastTransferTime,10=iTotalBytes,11=iRecipCountMax,12=iExpandCount,13=iDuplicateDeliverCount
			aFinalDeliveryDictionaryEntry(0) = 1									
			aFinalDeliveryDictionaryEntry(1) = aFinalLatencyTrackingEntry(1)
			aFinalDeliveryDictionaryEntry(2) = aFinalLatencyTrackingEntry(2)
			If aFinalLatencyTrackingEntry(4) <= iServerLatencySLA Then
				aFinalDeliveryDictionaryEntry(4) = 1
			Else
				aFinalDeliveryDictionaryEntry(4) = 0
			End If
			If aFinalLatencyTrackingEntry(7) Then 
				aFinalDeliveryDictionaryEntry(7) = 1			
			Else
				aFinalDeliveryDictionaryEntry(7) = 0
			End If
			aFinalDeliveryDictionaryEntry(8) = aFinalLatencyTrackingEntry(8)
			aFinalDeliveryDictionaryEntry(10) = aFinalLatencyTrackingEntry(10)
			aFinalDeliveryDictionaryEntry(11) = aFinalLatencyTrackingEntry(11)
			aFinalDeliveryDictionaryEntry(12) = aFinalLatencyTrackingEntry(12)
			aFinalDeliveryDictionaryEntry(13) = aFinalLatencyTrackingEntry(13)			
			oFinalDeliveryDictionary.Add sFinalDeliveryDictionaryKey, aFinalDeliveryDictionaryEntry			
		End If
	End If
	
	
End Sub

Sub CalculateQueueSize (ByVal sCurrentEventTime, byVal bSubmission, byVal iCurrentMessageSize)
	
	On Error Resume Next
	'For each message added or removed from queue database, calculate current queue size
	
	If bSubmission Then
		iCurrentQueueCount = oServerLatencyDictionary.Count + 1
		iCurrentQueueSizeBytes = iCurrentQueueSizeBytes + iCurrentMessageSize
	Else
		iCurrentQueueCount = oServerLatencyDictionary.Count - 1
		iCurrentQueueSizeBytes = iCurrentQueueSizeBytes - iCurrentMessageSize
	End If
	
	If iCurrentQueueCount > iMaxFileQueueCount Then
		iMaxFileQueueCount = iCurrentQueueCount
	End If
	If iCurrentQueueSizeBytes > iMaxQueueSizeBytes Then
		'Wscript.echo "iMaxQueueSizeBytes = " & iMaxQueueSizeBytes
		iMaxQueueSizeBytes = iCurrentQueueSizeBytes
	End If
	
	If iCurrentQueueCount >= iAggregateQueueSizeThreshold and sQueueExceptionStart = "01/01/1970 12:00:01 AM" Then
		sQueueExceptionStart = sCurrentEventTime
	ElseIf sQueueExceptionStart <> "01/01/1970 12:00:01 AM" Then
		iCurrentQueueExceptonDuration = datediff("s",sQueueExceptionStart,sCurrentEventTime)
		If iCurrentQueueExceptonDuration > 0 Then
			iQueueExceptonDurationSec = iQueueExceptonDurationSec + iCurrentQueueExceptonDuration
		End If
		sQueueExceptionStart = "01/01/1970 12:00:01 AM"
	End If
	
	If Second(sCurrentEventTime) Mod 15 = 0 Then 
		'Every 15 seconds, save queue size and rates of submission/delivery+send (reset variables used to track rates)
		'use dictionary to prevent creation of multiple data points for same time
		'dictionary should use server name as key with dictionary of data for each sample (key is time)
	Else
		'gather data to calculate rates
		'determine whether more than 15 seconds have passed since the last data point was saved (create data point if so)
	End If
	
End Sub


Sub WriteOrphanQueueEntry(byVal strQUniqueMessageId, byRef aLatencyTrackingEntry)
	'Wscript.echo "Orphan entry: " & strQUniqueMessageId & "," & aLatencyTrackingEntry(0) & "," & aLatencyTrackingEntry(1) & "," & aLatencyTrackingEntry(2) & "," & aLatencyTrackingEntry(3) & "," & aLatencyTrackingEntry(4) & "," & aLatencyTrackingEntry(5) & "," & aLatencyTrackingEntry(6) & "," & aLatencyTrackingEntry(7) & "," & aLatencyTrackingEntry(8) & "," & aLatencyTrackingEntry(9)
	If bWriteOrphanQueueEntries Then oQueueOrphanResults.writeline strQUniqueMessageId & "," & aLatencyTrackingEntry(0) & "," & aLatencyTrackingEntry(1) & "," & aLatencyTrackingEntry(2) & "," & aLatencyTrackingEntry(3) & "," & aLatencyTrackingEntry(4) & "," & aLatencyTrackingEntry(5) & "," & aLatencyTrackingEntry(6) & "," & aLatencyTrackingEntry(7) & "," & aLatencyTrackingEntry(8) & "," & aLatencyTrackingEntry(9) & "," & aLatencyTrackingEntry(10)
End Sub

Function GetOriginalArrivalTime(byVal sMessageInfo)
	On Error Resume Next
	If instr(1,sMessageInfo,";") Then
		aMessageInfo = split(sMessageInfo,";")
		strOriginalArrivalTime = GetDateTime(aMessageInfo(0))
	Else
		strOriginalArrivalTime = GetDateTime(sMessageInfo)
	End If
	GetOriginalArrivalTime = strOriginalArrivalTime
	'Wscript.echo sMessageInfo & vbtab & strOriginalArrivalTime
End Function