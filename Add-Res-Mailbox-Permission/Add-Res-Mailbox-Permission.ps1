# Powershell script to grant permission to resource mailboxes 
#  
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
# 
# Examples to use this script in Powershell... 
#  
# To add fullaccess permission for "administrator" account  
# 
#.\Add-Res-Mailbox-Permission "administrator" 
 
 
trap 
{ 
    write-host $_.Exception.Message -fore Red 
        continue 
} 
 
function Grand-FullPermission 
{ 
 
    $searcher = new-object DirectoryServices.DirectorySearcher 
 
    $searcher.filter = ("(&(objectCategory=person)(objectClass=User)(mailnickname=*)(userAccountControl:1.2.840.113556.1.4.803:=2))") 
 
    $user = $searcher.FindAll() 
 
 
    if ($user.Count -lt 1) 
    { 
    throw "no resource mailboxes found." 
    } 
    else 
    { 
 
        foreach($objResult in $user) 
        { 
      
            $objUser = $objResult.Properties 
 
            #Write-Host "Account Name (CN): $($objUser.name)" 
            #Write-Host "Alias (mailNickname): $($objUser.mailnickname)" 
 
            #add-content -path mailboxlist.txt -encoding ascii $($objUser.mailnickname)  
 
            Add-MailboxPermission -identity $($objUser.mailnickname) -user $script:serviceaccnt -accessrights fullaccess 
 
        } 
    } 
} 
 
 
$script:serviceaccnt = $args[0] 
if ($script:serviceaccnt.Length -gt 0) 
{ 
    Grand-FullPermission 
} 
else 
{ 
    write-host "no service account supplied" 
} 
