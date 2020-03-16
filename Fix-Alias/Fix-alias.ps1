# ================================================ 
# 
# This script allows you to identify and modify users, contact and distributions lists 
# that have invalid characters in the alis. 
#  
# Last Modified 05.03.07 
# Author: matbyrd@microsoft.com 
# 
# ================================================ 
 
# Input Parameters 
param ([string]$type = "",$resultsize = "Unlimited",[string]$search = " ",[string]$replace = "",$add = $null,[switch]$help=$false) 
 
# Variables 
 
[array]$baseobjects        # Array to hold all found objects 
[string]$new            # Value of the new alias 
$choice                # Used to determine choice in switch commands 
$command            # For holding the constructed command to execute 
 
# Formating Variables 
# Allows a Central place for defining the colors of script messages 
 
[string]$info = "White"        # Color for informational messages 
[string]$warning = "Yellow"    # Color for warning messages 
[string]$error = "Red"        # Color for error messages 
 
# ShowHelp Function (help about_function) 
# ================================================ 
 
Function ShowHelp { 
 
Write-host "This script will find objects of the specified type that contain a space in the alias" 
Write-host "It will remove the space from the alias and update the object" 
Write-host " " 
Write-host "Both the search character and the replacement character can be changed using the advanced options" 
Write-host "Advanced Options:" -foregroundcolor $warning 
write-host " " 
Write-host "-Type       : Used to specifiy the get- command that is run to find the objects (Mailbox,Distributiongroup,Mailcontact)" 
write-host "-Resultsize : Used to specifiy a result size other than the default of `"Unlimited`"" 
write-host "-Search     : Used to specify the character / sting to search for" 
write-host "-Replace    : Used to specify the replacement character" 
write-host "-Add        : Used to provide the get- command with addtional switch options" 
write-host "-Help       : Display this help message" 
write-host " " 
Write-host "Examples:" -foregroundcolor $info 
write-host " " 
Write-host "fix-alias.ps1 -type MailContact -Search `"@`" -Replace `"_`" -add `"-OrganizationalUnit 'My Ou'`"" 
 
} 
 
 
# Function to gather and modify the alias (help about_function) 
# ================================================ 
 
Function FixObject { 
 
# Use iex (invoke-expression) to execute the cmdlet in $command (help invoke-expresion) 
# Place the output in the $baseobjects array 
 
$baseobjects = iex $getcommand 
 
# Loop thru all of the object in the $baseobjects array (help about_foreach) 
# If the alias of the object contains the search character/characters operate on the object (help about_if) 
# Otherwise Do Nothing 
 
foreach ($value in $baseobjects) {  
    if ($value.alias -like "*$search*") 
        { 
            # Write out that we found an object to modify 
            # Use the string.replace .net method to search for and replace the character 
            # http://msdn2.microsoft.com/en-us/library/fk49wtc1(vs.90).aspx 
            # Write out the New Alias 
            # Construct the Set command into a variable and execute using iex 
             
            write-host "Found Object to Fix:" $value.alias -foregroundcolor $error 
             
            $new = $value.alias 
            $new = $new.replace($search,$replace) 
 
            write-host "New Alias of Object:" $new -foregroundcolor $info 
            write-host " " 
 
            $setcommand = "set-" + $type + " '" + $value.identity + "' -alias $new" 
            iex $setcommand 
             
        } 
    else { } 
} 
 
} 
 
# Main Body of Script 
# ================================================ 
 
# Display help if -help specified (help about_if) 
 
if ($help -eq $true) 
    { ShowHelp;exit } 
else { } 
 
# Determine if $type is set by parameter (help about_if) 
# Provide the user a list of choices if $type is not set 
 
if ( $type.length -eq 0 ) 
    { 
 
        # Provide user with choice of objects to check 
         
        Write-host " " 
        write-host "Please Choose what objects to search for" 
        write-host " " 
        write-host "1 - Mailbox" 
        write-host "2 - Contacts" 
        Write-host "3 - Distribution Group" 
         
        # Capture the users choice (help read-host) 
        $choice = read-host "Choice (1,2,3)" 
 
        # Using the switch command evaluate the input and set the $type variable (help about_switch) 
        # If an out of bounds choice is made show the script help 
 
        switch ($choice) { 
            1 {$type = "Mailbox"} 
            2 {$type = "MailContact"} 
            3 {$type = "distributiongroup"} 
            default {write-host " ";write-host "Incorrect options Specified: $choice" -foregroundcolor $error;ShowHelp;exit} 
        } 
             
    } 
else { } 
 
# Constuct the command to be executed into the $getcommand variable 
 
$getcommand = "get-" + $type + " -resultsize $resultsize " + $add 
 
# Call the FixObject function to fix the objects 
 
FixObject 
