#

ProcessTrackingLog.vbs

The Process Tracking Log tool simplifies parsing, monitoring, and analyzing Message Tracking logs by automating the parsing and analysis of Message tracking logs, and then reporting it back by producing xls or txt output files with meaningful data.

## Description

There are many scenarios in Exchange 2007 Server which requires parsing, monitoring, and analyzing Message Tracking logs. For example, in transaction log file and / or Database growth scenarios, Server got hit with spam messages, Looping message scenario, Transport queue backup scenarios, and Server performance scenarios, and many other administrative, monitoring, and troubleshooting scenarios - it absolutely becomes necessary to review and analyze Message Tracking Logs. The amount of fields and tremendous amount of data present in message tracking logs, and the amount of extremely high messaging traffic that Exchange servers process these days ( i.e. the sheer size of message tracking log files ) makes it extremely difficult, if not impossible, to analyze these message tracking logs manually. The task is further complicated when you have to review message tracing logs from multiple servers, and especially on a pretty regular basis.

Process Tracking Log tool simplifies these tasks by automating the parsing and analysis of Message tracking logs, and then reporting it back by producing xls or txt output files with meaningful data, which is extremely useful in a lot of Administration, troubleshooting, and monitoring scenarios in Exchange 2007, as discussed above. Besides rich useful data regarding monitoring, administration, and troubleshooting, the tool also provides critical data on End-To-End Delivery Latency Distribution for all Messages Delivered, Server Latency Distribution for all Messages Processed and Server Latency Distribution for all Individual Deliveries. The tool is developed by Todd Luttinen, Principal Program Manager at Microsoft, and is released with this blog post.

USAGE:

cscript ProcessTrackingLog.vbs <LogFilePath> <NumFiles> <hub|edge|all> [ <mm/dd/yyyy> | today | yesterday ]
For more information on this download see: http://blogs.technet.com/b/exchange/archive/2008/02/07/process-tracking-log-tool-for-exchange-server-2007.aspx

## Disclaimer

The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.