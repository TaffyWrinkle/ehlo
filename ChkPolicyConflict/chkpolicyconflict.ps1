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
 
 
# 1. Get list of recipient policies 
# 2. Loop through the email addresses in each policy 
# 3. Add each address to either the auth or non-auth array 
# 4. Compare each list and point out differences 
 
 
param([bool]$debugOutput=$false) 
$Error.Clear() 
 
function GetRecipientPolicies() 
{ 
    $Searcher.SearchScope = "Subtree" 
    $Searcher.Filter = "(objectClass=msExchRecipientPolicy)" 
    $Searcher.SearchRoot = $rootDN 
 
    $PolicyList = $Searcher.FindAll() 
    foreach($Policy in $PolicyList) 
    { 
        if($debug){write-host ("Processing Policy - " + $Policy.properties.cn) -ForeGroundColor Green} 
        AddPoliciesToHashTable($Policy) 
    } 
} 
 
 
function AddPoliciesToHashTable($Policy) 
{ 
    #Loop through each proxy address in gatewayProxy 
    ForEach($Proxy in $Policy.properties.gatewayproxy) 
    { 
 
        if($debug){write-host "    Processing Proxy" $Proxy.ToString().ToLower()} 
 
        #if the address is an SMTP address, check against non-authoritative policy list 
        if($Proxy.ToString().ToLower().StartsWith("smtp:")) 
        { 
            if($Policy.properties.msexchnonauthoritativedomains.Count -gt 0) 
            #non-authoritative policies exist on this policy object, add them to the non-authoritative list 
            { 
                #loop through each non-auth domain and check 
                ForEach($NonAuthDomain in $Policy.properties.msexchnonauthoritativedomains) 
                { 
                    #non-authoritative domain exists 
                    if ($Proxy.ToString().ToLower() -eq $NonAuthDomain.ToString().ToLower()) 
                    #add to non-authoritative domain list 
                    {     
                        if(-not $NonAuthoritativeDomainsList.Contains($Proxy.ToString() + "," + $Policy.properties.cn)) 
                        { 
                            #need to remove it from Auth list first since it is now a non-Auth domain 
                            if($AuthoritativeDomainsList.Contains($Proxy.ToString() + "," + $Policy.properties.cn)) 
                            { 
                                $AuthoritativeDomainsList.Remove($Proxy.ToString().ToLower() + "," + $Policy.properties.cn) 
                                if($debug){write-host ("        Removing " + $Proxy.ToString().ToLower() + " from Authoritative Proxy List") -ForeGroundColor Green} 
                            } 
                            if($debug){write-host ("        Adding " + $Proxy.ToString().ToLower() + " to Non-Authoritative Proxy List") -ForeGroundColor Green} 
                            $NonAuthoritativeDomainsList.Add($Proxy.ToString().ToLower() + "," + $Policy.properties.cn, $Policy.properties.cn) 
                             
                        } 
                        else 
                        #already in non-authoritative list 
                        { 
                            if($debug){write-host ("        Proxy " + $Proxy.ToString().ToLower() + " already exists in the Non-Authoritative Proxy List") -ForeGroundColor Yellow} 
                        } 
                    } 
                    else 
                    #authoritative policies exist as well on this policy, add them to authoritative list 
                    { 
                        if(-not $AuthoritativeDomainsList.Contains($Proxy.ToString().ToLower() + "," + $Policy.properties.cn) -and -not $NonAuthoritativeDomainsList.Contains($Proxy.ToString().ToLower() + "," + $Policy.properties.cn)) 
                        { 
                            if($debug){write-host ("        Adding " + $Proxy.ToString().ToLower() + " to Authoritative Proxy List") -ForeGroundColor Green} 
                            $AuthoritativeDomainsList.Add($Proxy.ToString().ToLower() + "," + $Policy.properties.cn, $Policy.properties.cn) 
                        } 
                        else 
                        #already in Authoritative list 
                        { 
                            if($debug){write-host ("        Proxy " + $Proxy.ToString().ToLower() + " already exists in the Authoritative Proxy List") -ForeGroundColor Yellow} 
                        } 
                         
                    }     
                } 
            } 
            else 
            #no non-authoritative policies exist, add smtp proxy to the authoritative list 
            { 
                if(-not $NonAuthoritativeDomainsList.Contains($Proxy.ToString() + "," + $Policy.properties.cn)) 
                { 
 
                if(-not $AuthoritativeDomainsList.Contains($Proxy.ToString() + "," + $Policy.properties.cn)) 
                { 
                    if($debug){write-host ("        Adding " + $Proxy.ToString().ToLower() + " to Authoritative Proxy List") -ForeGroundColor Green} 
                    $AuthoritativeDomainsList.Add($Proxy.ToString().ToLower() + "," + $Policy.properties.cn, $Policy.properties.cn) 
                } 
                else 
                #Item already exists in the Authoritative Proxy List 
                { 
                    if($debug){write-host ("        Proxy " + $Proxy.ToString().ToLower() + " already exists in the Authoritative Proxy List") -ForeGroundColor Yellow} 
                } 
                } 
            }     
        } 
        else 
        #non-SMTP proxy, log and continue 
        { 
            if($debug){write-host ("        Skipping " + $Proxy.ToString() + " - Proxy is not an SMTP Proxy type") -ForeGroundColor Yellow} 
        } 
         
    } 
} 
 
function WriteOutput() 
{ 
    write-host  
    write-host 
 
    #write authoritative domain results 
    write-host ("Total Authoritative Domains - " + $AuthoritativeDomainsList.Count) -ForeGroundColor Green 
    ForEach($item in $AuthoritativeDomainsList.Keys) 
    { 
         
        write-host "   Domain:" $item.Split(',')[0] " Policy:" $AuthoritativeDomainsList.Item($item) 
    } 
 
    #write non-authoritative domain results 
    write-host ("Total Non-Authoritative Domains - " + $NonAuthoritativeDomainsList.Count) -ForeGroundColor Green 
    ForEach($item in $NonAuthoritativeDomainsList.Keys) 
    {     
         
        write-host "   Domain:" $item.Split(',')[0] " Policy:" $NonAuthoritativeDomainsList.Item($item) 
    } 
} 
 
function CompareLists() 
{ 
 
$errorOutput 
$errorAuth 
$errorCount = 0         
 
    ForEach($item in $AuthoritativeDomainsList.Keys) 
    { 
        ForEach($nonauthitem in $NonAuthoritativeDomainsList.Keys) 
        { 
            if($item.ToString().Split(',')[0].ToLower() -eq $nonauthitem.ToString().Split(',')[0].ToLower()) 
            { 
                $errorAuth =  "   Authoritative in policy " + $AuthoritativeDomainsList.Item($item) 
                $errorOutput +=  "   Non-Authoritative in policy " + $NonAuthoritativeDomainsList.Item($nonauthitem) + "`r`n" 
                #write-host  
                #write-host ("Found conflicting domain " + $item.ToString().Split(',')[0].ToLower()) -ForeGroundColor Red 
                #write-host ("   Authoritative in policy " + $AuthoritativeDomainsList.Item($item)) -ForeGroundColor Red 
                #write-host ("   Non-Authoritative in policy " + $NonAuthoritativeDomainsList.Item($nonauthitem)) -ForeGroundColor Red 
                 
            } 
        } 
        if($errorOutput.Length -gt 0) 
        { 
            write-host ("`r`nFound conflicting domain " + $item.ToString().Split(',')[0].ToLower()) -ForeGroundColor Red 
            write-host $errorAuth -ForeGroundColor Red 
            write-host $errorOutput -ForeGroundColor Red 
            $errorCount++ 
        } 
        $errorOutput = $null 
    } 
 
    if($errorCount -eq 0) 
    { 
        write-host ("`r`nNo conflicting domains found") -ForeGroundColor Green 
    } 
         
} 
 
 
if($debugOutput){$debug=$true} 
 
$AuthoritativeDomainsList = @{} 
$NonAuthoritativeDomainsList = @{} 
 
$Searcher = New-Object DirectoryServices.DirectorySearcher 
$Root = New-Object DirectoryServices.DirectoryEntry("GC://rootDSE") 
$rootDN = New-Object DirectoryServices.DirectoryEntry("LDAP://" + $Root.configurationNamingContext) 
 
 
GetRecipientPolicies 
WriteOutput 
CompareLists
