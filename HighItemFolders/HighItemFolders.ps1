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
 
 
param( 
[bool] $CriticalFoldersOnly = $true, 
[string] $Database = "", 
[string] $DomainController = "", 
[bool] $FormatList = $false, 
[int32] $ItemCount = -1, 
[string] $OutputFile = "", 
[string] $ResultSize = "unlimited", 
[string] $Server = "") 
 
#Script Parameters: 
#-CriticalFoldersOnly: Only checks Calendar, Contacts, Inbox, and Sent Items. The default value is $true. 
#-Database: Specifies the target Database. Overrides the -Server option if set. 
#-DomainController: Specifies the default Domain Controller to use for mailbox and folder tests. 
#-FormatList: Outputs to the screen in list format. The default is table format. 
#-ItemCount: Ignores the max limits, and finds folders with the specified item count. 
#-OutputFile: Specifies the output file. Should be in .CSV format. 
#-ResultSize: Specifies the maximum number of mailboxes to check. The default value is unlimited. 
#-Server: Specifies the target Exchange server. 
 
# Function used to check the individual item count of folders within a given mailbox 
function checkFolderItems 
{ 
    param( 
    $mbx, 
    [ref]$fldrArray, 
    [ref]$mbxFailed) 
 
    #Set item limit to search for 
    [int32] $maxItems = 5000 
 
    if ($mbx.ExchangeVersion -eq "2000") 
    { 
        $maxItems = $2kmax 
    } 
    elseif ($mbx.ExchangeVersion -eq "2003") 
    { 
        $maxItems = $2k3max 
    } 
    elseif ($mbx.ExchangeVersion -eq "2007") 
    { 
        $maxItems = $2k7max 
    } 
    elseif ($mbx.ExchangeVersion -eq "2010") 
    { 
        $maxItems = $2k10max 
    } 
     
    #Create the base of our Get-MailboxFolderStatistics command 
    $getFoldersString = "Get-MailboxFolderStatistics -Identity `"$($mbx.Identity)`" -ErrorAction SilentlyContinue" 
 
    #Add the domain controller if specified 
    if ($DomainController -ne "") 
    { 
        $getFoldersString += " -DomainController $DomainController" 
    } 
 
    #Add the filter portion of the command 
    if ($CriticalFoldersOnly -eq $true) 
    { 
        $getFoldersString += " | Where {(`$_.Name -like 'Calendar' -or `$_.Name -like 'Contacts' -or `$_.Name -like 'Inbox' -or `$_.Name -like 'Sent Items') -and `$_.ItemsInFolder -ge $maxItems}" 
    } 
    else 
    { 
        $getFoldersString += " | Where {`$_.ItemsInFolder -ge $maxItems}" 
    } 
 
    #Get the current error count so we can report errors after checking folders 
    $errorCountBefore = $error.Count 
     
    #Actually run our command and process all folders over the item limit 
    Invoke-Expression -Command $getFoldersString | ForEach-Object{ 
        $folder = New-Object PSObject 
 
        $folder | Add-Member NoteProperty User $mbx.Alias -Force 
        $folder | Add-Member NoteProperty Folder $_.FolderPath -Force 
        $folder | Add-Member NoteProperty ItemCount $_.ItemsInFolder -Force 
        $folder | Add-Member NoteProperty Database $mbx.Database -Force 
        $folder | Add-Member NoteProperty Version $mbx.ExchangeVersion -Force 
 
        if ($fldrArray.Value[0] -ne "") 
        { 
            $fldrArray.Value += @($folder) 
        } 
        else 
        { 
            $fldrArray.Value = @($folder) 
        } 
    } 
 
    #Now check the current error count. Report an error if it is larger than when we started 
    if ($error.Count -gt $errorCountBefore) 
    { 
        Write-Host -ForeGroundColor red "ERROR: Unable to process mailbox '$($mbx.Alias)'." 
        $mbxFailed.Value++ 
    } 
} 
 
# Function that checks whether the target server or database is inaccessible. 
# This prevents us from timing out on each inaccessible mailbox. 
# Note, this does not work against 2000-2003 mailboxes. 
# Returns $true if accessible, and $false is not. 
function checkAccessibility 
{ 
    param( 
    [ref]$badServers, 
    [ref]$badDBs, 
    [ref]$goodServers, 
    [ref]$goodDBs, 
    [ref]$mbxFailed, 
    $mbx) 
 
 
    #Check database and server accessibility. If mailbox is not on Exchange 2007+, or the server or database are already bad, fail. 
    if ($mbx.ExchangeVersion -ge 2007 -and !($badDBs.Value -Contains $mbx.Database) -and !($badServers.Value -Contains $mbx.ServerName)) 
    { 
        #Check if mailbox is on a higher version of Exchange than this shell 
        if ($mbx.ExchangeVersion -ge 2010 -and (Get-Command "Microsoft.Exchange.PowerShell.Configuration.dll").FileVersionInfo.FileVersion -lt "14") 
        { 
            Write-Host -ForeGroundColor red "ERROR: Server '$($mbx.ServerName)' is running a higher version of Exchange than this machine. Skipping all mailboxes on this server." 
 
            $badServers.Value += $mbx.ServerName 
            $mbxFailed.Value++ 
            return $false         
        } 
         
        #Check if database or server is already marked as good. 
        $serverGood = $false 
        $dbGood = $false 
         
        if ($goodServers.Value -Contains $mbx.ServerName) 
        { 
            $serverGood = $true 
        } 
         
        if ($goodDBs.Value -Contains $mbx.Database) 
        { 
            $dbGood = $true 
        } 
 
        #Either the server or database have not been marked as good. Proceed with connectivity test. 
        if (!$serverGood -or !$dbGood) 
        { 
            $failed = $false 
 
            $testMAPIString = "Test-MAPIConnectivity -Database '$($mbx.Database)' -PerConnectionTimeout 10 -ErrorAction SilentlyContinue" 
 
            if ($DomainController -ne "") 
            { 
                $testMAPIString += " -DomainController $DomainController" 
            } 
 
            $testMAPI = Invoke-Expression -Command $testMAPIString 
 
            #Check if the server is accessible 
            if (!$serverGood -and $testMAPI.Error -like "Microsoft Exchange Information Store service is not running.") 
            { 
                Write-Host -ForeGroundColor red "ERROR: Server '$($mbx.ServerName)' is inaccessible. Skipping all mailboxes on this server." 
 
                $badServers.Value += $mbx.ServerName 
                $failed = $true                     
            } 
            else 
            { 
                $goodServers.Value += $mbx.ServerName 
            } 
              
            #Check if the database is accessible 
            if (!$dbGood -and $testMAPI.Result -like "*Failure*" -and $testMAPI.Error -notlike "Microsoft Exchange Information Store service is not running.") 
            { 
                Write-Host -ForeGroundColor red "ERROR: Database '$($mbx.Database)' is inaccessible. Skipping all mailboxes on this database."              
 
                $badDBs.Value += $mbx.Database 
                $failed = $true 
            } 
            else 
            { 
                $goodDBs.Value += $mbx.Database 
            } 
              
            #Check if we failed 
            if ($failed) 
            { 
                $mbxFailed.Value++ 
                return $false 
            } 
            else 
            { 
                return $true 
            } 
        } 
        else 
        { 
            #The server and database are both in the good list. Proceed with checking mailbox 
            return $true 
        } 
    } 
    elseif ($mbx.ExchangeVersion -lt 2007) 
    { 
        #Do nothing and proceed to checking mailbox 
        return $true 
    } 
    else 
    { 
        #The server or database is in the bad list. Skip this mailbox 
        $mbxFailed.Value++ 
        return $false 
    } 
} 
 
# Function that writes output to the screen, and writes to the output file if specified 
function writeOutput 
{ 
    param( 
    [ref]$fldrArray) 
 
    [int32] $maxItems = 5000 
 
    if ($fldrArray.Value[0].Version -eq "2000") 
    { 
        $maxItems = $2kmax 
    } 
    elseif ($fldrArray.Value[0].Version -eq "2003") 
    { 
        $maxItems = $2k3max 
    } 
    elseif ($fldrArray.Value[0].Version -eq "2007") 
    { 
        $maxItems = $2k7max 
    } 
    elseif ($fldrArray.Value[0].Version -eq "2010") 
    { 
        $maxItems = $2k10max 
    } 
 
    #Write output to screen 
    $screenOutput = @() 
 
    Write-Host "" 
    Write-Host -ForeGroundColor green "Exchange $($fldrArray.Value[0].Version) Folders With $($maxItems) or More Items:"         
 
    for ($i = 0; $i -lt $fldrArray.Value.Length; $i++) 
    { 
        $screenOutput += $($fldrArray.Value[$i]) 
 
        #Write output to file 
        if ($usingOutput -eq $true) 
        { 
            $currentLine = "$($fldrArray.Value[$i].User)`t$($fldrArray.Value[$i].Folder)`t$($fldrArray.Value[$i].ItemCount)`t$($fldrArray.Value[$i].Database)`t$($fldrArray.Value[$i].Version)" 
            $currentLine | Out-File -Append -NoClobber $OutputFile     
        } 
    } 
 
    #Check whether we should output in list or table format 
    if ($FormatList) 
    { 
        Write-Output $screenOutput | fl User, Folder, ItemCount, Database 
    } 
    else 
    { 
        Write-Output $screenOutput | ft User, Folder, ItemCount, Database -Autosize 
    } 
 
    #Add the folders from this pass to the master output array 
    $outputArray += $screenOutput 
    $screenOutput = $null 
} 
 
#Function that parses the ExchangeVersion property on a mailbox 
function getExchangeVersion 
{ 
    param( 
    [ref]$mbx) 
     
    [double] $mbxVer = 0 
 
    #Check whether ExchangeVersion is set properly on the mailbox 
    if ($mbx.Value.ExchangeVersion.ExchangeBuild.Major -lt 6 -or $mbx.Value.ExchangeVersion.ExchangeBuild.Major -gt 20) 
    { 
        #ExchangeVersion is not set properly. Defaulting to Exchange 2010 
        $mbxVer = 14 
    } 
    else 
    { 
        $mbxVer = $mbx.Value.ExchangeVersion.ExchangeBuild.Major + ($mbx.Value.ExchangeVersion.ExchangeBuild.Minor/10) 
    } 
 
    #Convert the version to a string 
    #The mailbox is 2010 
    if ($mbxVer -ge 14) 
    { 
        $mbxVer = "2010" 
    } 
    #The mailbox is 2007 
    elseif ($mbxVer -ge 8 -and $mbxVer -le 8.5) 
    { 
        $mbxVer = "2007" 
    } 
    #The mailbox is 2003 
    elseif ($mbxVer -eq 6.5) 
    { 
        $mbxVer = "2003" 
    } 
    #The mailbox is 2000 
    elseif ($mbxVer -ge 6 -and $mbxVer -lt 6.5) 
    { 
        $mbxVer = "2000" 
    } 
    #We didn't find the version. Default to 2010 
    else 
    { 
        $mbxVer = "2010" 
    } 
 
    return $mbxVer 
} 
 
# Function that returns true if the incoming argument is a help request 
function IsHelpRequest 
{ 
    param($argument) 
    return ($argument -eq "-?" -or $argument -eq "-help"); 
} 
 
# Function that displays the help related to this script following 
# the same format provided by get-help or <cmdletcall> -? 
function Usage 
{ 
@" 
 
NAME: 
`tHighItemFolders.ps1 
 
SYNOPSIS: 
`tFinds users who have folders with more than the recommended 
`titem count. The item count that is searched for is different for 
`teach Exchange version: 
 
`t`tExchange 2000: 5000 Items 
`t`tExchange 2003: 5000 Items 
`t`tExchange 2007: 20,000 Items 
`t`tExchange 2010: 100,000 Items 
 
SYNTAX: 
`tHighItemFolders.ps1 
`t`t[-CriticalFoldersOnly <BooleanValue>] 
`t`t[-Database <DatabaseIdParameter>] 
`t`t[-DomainController <StringValue>] 
`t`t[-FormatList <BooleanValue>] 
`t`t[-ItemCount <IntegerValue>] 
`t`t[-OutputFile <OutputFileName>] 
`t`t[-ResultSize <StringValue>] 
`t`t[-Server <ServerIdParameter>] 
 
PARAMETERS: 
`t-CriticalFoldersOnly 
`t`tSpecifies whether to check only Critical Folders, which 
`t`tare Calendar, Contacts, Inbox, and Sent Items. Should be 
`t`tinput as either `$true or `$false. If omitted, the default 
`t`tvalue is $true. 
 
`t-Database 
`t`tSpecifies the target database. 
`t`tOverrides the -Server switch if used. 
 
`t-DomainController 
`t`tSpecifies the Domain Controller to use for all mailbox 
`t`tand folder tests. If omitted, the default value is `$null. 
 
`t-FormatList 
`t`tWrites output to the screen in list format instead of 
`t`ttable format. Should be input as either `$true or `$false. 
`t`tIf omitted, the default value is `$false. 
 
`t-ItemCount 
`t`tIgnores the max limits, and finds folders with the 
`t`tspecified item count. 
 
`t-OutputFile 
`t`tSpecifies the output file. Should be in .CSV format. 
 
`t-ResultSize 
`t`tSpecifies the maximum number of mailboxes to be checked. 
`t`tCan be specified either as a number, or 'Unlimited'. 
`t`tIf omitted, the default value is unlimited. 
 
`t-Server 
`t`tSpecifies the target Exchange server. 
 
`t-------------------------- EXAMPLES ---------------------------- 
 
C:\PS> .\HighItemsUsers.ps1 -Server "MyEx2007Server" -OutputFile output.csv -CriticalFoldersOnly `$false 
 
C:\PS> .\HighItemsUsers.ps1 -Database "MyEx2007Server\My Storage Group\My Database" 
 
C:\PS> .\HighItemsUsers.ps1 
 
REMARKS: 
`tIf -Database and -Server are omitted, the entire Organization 
`twill be checked. 
 
"@ 
} 
 
#################################################################################################### 
# Script starts here 
#################################################################################################### 
 
# Check for Usage Statement Request 
$args | foreach { if (IsHelpRequest $_) { Usage; exit; } } 
 
#Declare the arrays for holding problem folders. 
#The arrays are initialized to "" so that they can be passed by 
#reference, even if they contain no objects. 
[Array] $2kFolders = "" 
[Array] $2k3Folders = "" 
[Array] $2k7Folders = "" 
[Array] $2k10Folders = "" 
 
#Declare the arrays for keeping track of inaccessible servers and databases. 
[Array] $badServers = @() 
[Array] $badDBs = @() 
[Array] $goodServers = @() 
[Array] $goodDBs = @() 
 
#Declare the item count limits we are looking for 
if ($ItemCount -gt -1) 
{ 
    $2kmax = $ItemCount 
    $2k3max = $ItemCount 
    $2k7max = $ItemCount 
    $2k10max = $ItemCount 
} 
else 
{ 
    $2kmax = 5000 
    $2k3max = 5000 
    $2k7max = 20000 
    $2k10max = 100000 
} 
 
#Declare variables for keeping track of progress 
[int32]$mbxSucceeded = 0 
[int32]$mbxCount = $mbxCmd.Count 
[int32]$mbxFailed = 0 
 
#Determine which Get-Mailbox command to run 
$getMbxString = "Get-Mailbox -ResultSize $ResultSize" 
 
if ($Server -ne "") 
{ 
    $getMbxString += " -Server $Server" 
} 
elseif ($Database -ne "") 
{ 
    $getMbxString += " -Database '$Database'" 
} 
else 
{ 
    #Do nothing. Use the default command 
} 
 
if ($DomainController -ne "") 
{ 
    $getMbxString += " -DomainController $DomainController" 
} 
 
#This array will hold just the properties we need for each mailbox. This helps to conserve memory 
$mailboxes = @() 
 
#Start our progress bar 
Write-Progress -Activity "Getting All Mailboxes in the Specified Scope" -Status "Command: $getMbxString" 
 
#Actually run the command and store the output in $mailboxes 
Invoke-Expression -Command $getMbxString | ForEach-Object{ 
    $ver = getExchangeVersion -mbx ([ref]$_) 
 
    $mbx = New-Object PSObject 
 
    $mbx | Add-Member NoteProperty Identity $_.Identity.ToString() 
    $mbx | Add-Member NoteProperty Alias $_.Alias.ToString() 
    $mbx | Add-Member NoteProperty ServerName $_.ServerName.ToString() 
    $mbx | Add-Member NoteProperty Database $_.Database.ToString() 
    $mbx | Add-Member NoteProperty ExchangeVersion $ver 
 
    $mailboxes += $mbx 
} 
 
#Close progress bar 
Write-Progress -Activity "Getting All Mailboxes in the Specified Scope" -Completed -Status "Completed" 
 
#Get total mailbox count 
$mbxCount = $mailboxes.Count 
 
#Begin processing of mailboxes 
foreach ($mailbox in $mailboxes) 
{ 
    #Update Progress Bar 
    $failCountBefore = $mbxFailed 
    Write-Progress -Activity "Checking $mbxCount Mailboxes for High Item Counts" -Status "Mailboxes Successfully Processed: $mbxSucceeded   Inaccessible Mailboxes: $mbxFailed" 
 
    #Check whether the server and database of this mailbox are accessible (only works on 2007 and higher) 
    if (!(checkAccessibility -badServers ([ref]$badServers) -badDBs ([ref]$badDBs) -goodServers ([ref]$goodServers) -goodDBs ([ref]$goodDBs) -mbxFailed ([ref]$mbxFailed) -mbx $mailbox)) 
    { 
        #We failed the accessibility check. Skip this mailbox. 
        Continue 
    } 
 
    #Now on to the actual item checking 
    #Mailbox version is Exchange 2000 
    if ($mailbox.ExchangeVersion -eq "2000") 
    { 
        checkFolderItems -mbx $mailbox -fldrArray ([ref]$2kFolders) -mbxFailed ([ref]$mbxFailed) 
    } 
 
    #Mailbox version is Exchange 2003 
    elseif ($mailbox.ExchangeVersion -eq "2003") 
    { 
        checkFolderItems -mbx $mailbox -fldrArray ([ref]$2k3Folders) -mbxFailed ([ref]$mbxFailed) 
    } 
 
    #Mailbox version is Exchange 2007 
    elseif ($mailbox.ExchangeVersion -eq "2007") 
    { 
        checkFolderItems -mbx $mailbox -fldrArray ([ref]$2k7Folders) -mbxFailed ([ref]$mbxFailed) 
    }     
 
    #Mailbox version is Exchange 2010 
    elseif ($mailbox.ExchangeVersion -eq "2010") 
    { 
        checkFolderItems -mbx $mailbox -fldrArray ([ref]$2k10Folders) -mbxFailed ([ref]$mbxFailed) 
    } 
 
    if ($failCountBefore -eq $mbxFailed) 
    { 
        $mbxSucceeded++ 
    } 
} 
 
#Done processing mailboxes. Stop the progress bar 
Write-Progress -Activity "Checking $mbxCount Mailboxes for High Item Counts" -Completed -Status "Completed" 
 
#Initialize the Output File 
$usingOutput = $false 
 
if ($OutputFile -ne "") 
{ 
    $usingOutput = $true 
    "Alias`tFolderPath`tItemsInFolder`tDatabase`tExchangeVersion" | Out-File $OutputFile 
} 
 
#Sort all the folders based on item count 
Write-Progress -Activity "Sorting Folders by Item Count" -Status " " 
 
$2kFolders = $2kFolders | Sort-Object -Property ItemCount -Descending 
$2k3Folders = $2k3Folders | Sort-Object -Property ItemCount -Descending 
$2k7Folders = $2k7Folders | Sort-Object -Property ItemCount -Descending 
$2k10Folders = $2k10Folders | Sort-Object -Property ItemCount -Descending 
 
Write-Progress -Activity "Sorting Folders by Item Count" -Completed -Status "Completed" 
 
#Start writing folder output 
$totalFolders = 0 
 
#Write Exchange 2000 Folders to Screen and Output File 
if ($2kFolders[0] -ne "") 
{ 
    writeOutput ([ref]$2kFolders) 
    $totalFolders += $2kFolders.Length 
} 
 
#Write Exchange 2003 Folders to Screen and Output File 
if ($2k3Folders[0] -ne "") 
{ 
    writeOutput ([ref]$2k3Folders) 
    $totalFolders += $2k3Folders.Length 
} 
 
#Write Exchange 2007 Folders to Screen and Output File 
if ($2k7Folders[0] -ne "") 
{ 
    writeOutput ([ref]$2k7Folders) 
    $totalFolders += $2k7Folders.Length 
} 
 
#Write Exchange 2010 Folders to Screen and Output File 
if ($2k10Folders[0] -ne "") 
{ 
    writeOutput ([ref]$2k10Folders) 
    $totalFolders += $2k10Folders.Length 
} 
 
#Write final statistics to screen 
Write-Host "" 
Write-Host -ForeGroundColor Green "Finished Processing Mailboxes" 
Write-Host -ForeGroundColor yellow "Total Mailboxes Found: " -NoNewLine 
Write-Host "$mbxCount" 
Write-Host -ForeGroundColor yellow "Mailboxes Successfully Processed: " -NoNewLine 
Write-Host "$mbxSucceeded" 
Write-Host -ForeGroundColor yellow "Mailboxes Skipped: " -NoNewLine 
Write-Host "$mbxFailed" 
Write-Host -ForeGroundColor yellow "High Item Folders Found: " -NoNewLine 
Write-Host "$totalFolders"
