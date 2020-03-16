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
 
 
param([string]$SearchRoot,[string]$Filter,[string]$Scope,[string]$Attribute,[string]$Value) 
 
$Error.Clear() 
 
########GLOBAL TRAP HANDLER######## 
 
trap [System.Management.Automation.MethodInvocationException] 
{ 
    write-host ("ERROR: " + $_.Exception) -ForeGroundColor Red; Continue 
} 
 
 
########OBJECT INSTANTIATION######## 
 
$Searcher = new-object DirectoryServices.DirectorySearcher 
$Root = [ADSI]("LDAP://" + $SearchRoot) 
 
 
########FUNCTION DEFINITIONS######## 
 
function IsHelpRequest 
{ 
    param($argument) 
    return($argument -eq "-?" -or $argument -eq "-help"); 
} 
 
function Usage 
{ 
@" 
 
NAME: 
ADModify.ps1 
 
SYNOPSIS: 
The purpose of this script is for bulk modification of Active Directory  
Objects.  All arguments are required for the script to run properly,  
below is a list of what these arguments mean: 
 
SearchRoot - An Active Directory path that points to the container in  
which to start the search.  An example would be "CN=Users,DC=Domain,DC=com" 
This value must be enclosed in double quotes. 
 
Filter - A valid LDAP filter to apply when searching for objects.   
An example would be "(&(objectClass=user)(description=Manager))"   
This value must be enclosed in double quotes. 
 
Scope - Specifies whether the search should be a Base, OneLevel, or  
Subtree search. 
 
Attribute - Specifies the name of the attribute that you wish to modify. 
 
Value - The value you want to set for the Attribute specified.  If you 
want to set an attribute to have the same value as another AD attribute 
the value should be enclosed in percent signs. For example, say you want 
to set everyones displayName to match their cn.  In this case Value would 
be %cn%. 
 
SYNTAX: 
ADModify.ps1 SearchRoot Filter Scope Attribute Value 
 
EXAMPLES: 
ADModify.ps1 "CN=Users,DC=Domain,DC=com" "(&(objectClass=user)(description=Manager)) Subtree description %cn% 
ADModify.ps1 "CN=Users,DC=Domain,DC=com" "(&(objectClass=user)(description=Manager)) Subtree extensionAttribute1 "Development" 
 
"@ 
 
} 
 
function SetSearchOptions() 
{ 
    $Searcher.SearchScope = $Scope 
    $Searcher.Filter = $Filter 
    $Searcher.SearchRoot = $Root 
} 
 
function GetSearchResults($UserList) 
{ 
    #Loop through the results and run modification on each one 
    foreach($User in $UserList) 
    { 
        write-host $User.Path 
        ModifyUser($User) 
    } 
    return 
} 
 
function ModifyLiteral($UserADSI) 
{ 
    #Trap handler needs to be inside the scope of this function or it will not continue if 
    #an exception is thrown 
    trap [System.Runtime.InteropServices.COMException] 
    { 
        if($_.Exception -match "The directory property cannot be found in the cache.") 
        { 
            write-host $Attribute "is blank for this user" -ForeGroundColor Yellow; Return $NewValue = "GoToNextUser" 
        } 
        else 
        { 
            write-host "WARNING:" $_.Exception -ForeGroundColor Yellow; Continue 
        } 
    } 
 
    $CurrentValue = $UserADSI.Get($Attribute) 
    write-host "Current Value is" $CurrentValue -ForeGroundColor White 
    $Put = $UserADSI.Put($Attribute,$Value) 
    $Set = $UserADSI.SetInfo() 
    $NewValue = $UserADSI.Get($Attribute) 
    return $NewValue 
} 
 
function ModifyVariable($UserADSI) 
{ 
     
    #This denotes whether or not the user entered another attribute to be resolved for 
    #the new attributes value.  Mostly for error handling purposes. 
    $AreWeInsideValueResolution = $false 
 
    #Trap handler needs to be inside the scope of this function or it will not continue if 
    #an exception is thrown 
    trap [System.Runtime.InteropServices.COMException] 
    { 
        if($_.Exception -match "The directory property cannot be found in the cache.") 
        { 
            if($AreWeInsideValueResolution) 
            { 
                write-host "Resolving attribute" $Value.Replace("%","") "failed" -ForeGroundColor Red; Return $NewValue = "GoToNextUser" 
            } 
            else 
            { 
                write-host $Attribute "is blank for this user" -ForeGroundColor Yellow; Continue 
            } 
        } 
        else 
        { 
                write-host "WARNING:" $_.Exception -ForeGroundColor Yellow; Continue 
        } 
    } 
 
    $CurrentValue = $UserADSI.Get($Attribute) 
    if($CurrentValue) 
    { 
        write-host "Current Value is" $CurrentValue -ForeGroundColor White 
    } 
    $AreWeInsideValueResolution = $true 
    $Value = $UserADSI.Get($Value.Replace("%","")) 
    $Put = $UserADSI.Put($Attribute, $Value) 
    $Set = $UserADSI.SetInfo() 
    $NewValue = $UserADSI.Get($Attribute) 
    $AreWeInsideValueResolution = $false 
    Return $NewValue,$Value 
} 
 
function ModifyUser($User) 
{ 
 
    $UserADSI = [ADSI]$User.Path 
 
    #If value is contained in percent signs treat it an AD attribute value instead of a literal value 
    if($Value.StartsWith("%") -and $Value.EndsWith("%")) 
    { 
        $NewValue = ModifyVariable($UserADSI) 
        $Value = $NewValue[1] 
        $NewValue = $NewValue[0] 
    } 
    #Treat as a literal value 
    else 
    { 
        $NewValue = ModifyLiteral($UserADSI) 
    } 
    #Check for failed attribute resolution, if failed go to next user 
    if($NewValue[0] -eq "GoToNextUser") 
    { 
        return 
    } 
 
 
    #Check if modification succeeded 
    if($Value -eq $NewValue) 
    { 
        write-host "New value is" $Value -ForeGroundColor Green 
    } 
    else 
    { 
        write-host "Attribute change failed" -ForeGroundColor Red 
    } 
    return 
} 
 
########BEGIN SCRIPT######## 
 
#Check for help request 
$args | foreach {if (IsHelpRequest $_) {Usage; exit; } } 
 
#Verify arguments 
if(!$SearchRoot -or !$Filter -or !$Scope -or !$Attribute -or !$Value) 
{ 
    write-host "One or more required arguments are missing" 
    break 
} 
 
#Begin Processing 
SetSearchOptions 
$UserList = $Searcher.FindAll() 
GetSearchResults($UserList) 
