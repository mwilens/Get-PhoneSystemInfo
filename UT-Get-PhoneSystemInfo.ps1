[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

Function UT_Display-Results{
  [CmdletBinding()]
  Param
  ( 
    [Parameter(Mandatory=$False)]
    [string]$BlockText,   # like info
    [string]$CheckName,   #
    [string]$TestColor,   # 1 of 16 available colors
    [object]$TestDisplay, #
    [object]$TestTable,   # @{Property="Name","Name2"}
    [object]$ShowResults, # InfoOnly, DataAndInfo, DataOnly(=without BlockText="info"), All
    [int]$MaxResults=50,
    [int]$MaxWith   =120
    )
# $MailBody is defined outside this function, what about : return MailBody
$BlockText = ($BlockText + "      ").Substring(0,6).Trim()
$BlockText = ("[" + $BlockText + "]" + "        ").Substring(0,8)
$TestDate  = get-date -Format HH:mm:ss
if ($TestDisplay -ne $null -and ($TestDisplay).count -eq 0 -or (!$TestDisplay)){$TestCount = ""}
if ($TestDisplay -ne $null -and ($TestDisplay).count -gt 0){$TestCount = ": " + ($TestDisplay).count}
if (($TestColor   -eq "") -or ($TestColor   -eq $null)){$TestColor = "Red"}

If ($ShowResults -like "All"          ){
    Write-Host "$BlockText $TestDate " -ForeGroundColor Yellow -NoNewLine
    Write-Host "$CheckName $TestCount" -ForegroundColor White
    $script:MailBody += "$BlockText $TestDate $CheckName $TestCount`r`n"

    if ($TestDisplay -ne $null -and (($TestDisplay).count -le $MaxResults) -and (($TestDisplay).count -ne 0)){
    $TestDisplay | Format-Table @TestTable -AutoSize | Out-String -Width 120 | Write-Host -ForegroundColor $TestColor -NoNewLine
    $script:MailBody += $TestDisplay | Format-Table @TestTable -AutoSize | Out-String -Width 120}
    
    if ($TestDisplay -ne $null -and ($TestDisplay).count -gt $MaxResults){
        Write-Host "                  ...Too many results, not showing..." -ForegroundColor $TestColor
        Write-Host
        $script:MailBody += "                  ...Too many results, not showing...`r`n`r`n"}
     }

If (($ShowResults -like "DataOnly") -and ($TestDisplay -ne $null)){
    Write-Host "$BlockText $TestDate " -ForeGroundColor Yellow -NoNewLine
    Write-Host "$CheckName $TestCount" -ForegroundColor White
    $script:MailBody += "$BlockText $TestDate $CheckName $TestCount`r`n"

    if (($TestDisplay).count -le $MaxResults){
    $TestDisplay | Format-Table @TestTable -AutoSize | Out-String -Width 120| Write-Host -ForegroundColor $TestColor -NoNewLine
    $script:MailBody += $TestDisplay | Format-Table @TestTable -AutoSize | Out-String -Width 120}
    
    if (($TestDisplay).count -gt $MaxResults){
       Write-Host "                  ...Too many results, not showing..." -ForegroundColor $TestColor
       Write-Host
       $script:MailBody += "                  ...Too many results, not showing...`r`n`r`n"}
    }

If (($ShowResults -like "DataAndInfo") -and (($TestDisplay -ne $null) -or ($BlockText -like "*Info*"))) {
    Write-Host "$BlockText $TestDate " -ForeGroundColor Yellow -NoNewLine
    Write-Host "$CheckName $TestCount" -ForegroundColor White
    $script:MailBody += "$BlockText $TestDate $CheckName $TestCount`r`n"

    if (($TestDisplay -ne $null) -and ($TestDisplay).count -le $MaxResults){
    $TestDisplay | Format-Table @TestTable -AutoSize | Out-String -Width 120 | Write-Host -ForegroundColor $TestColor -NoNewLine
    $script:MailBody += $TestDisplay | Format-Table @TestTable -AutoSize | Out-String -Width 120}
    
    if ($TestDisplay -ne $null){
       if(($TestDisplay).count -gt $MaxResults){
       Write-Host "                  ...Too many results, not showing..." -ForegroundColor $TestColor
       Write-Host
       $script:MailBody += "                  ...Too many results, not showing...`r`n`r`n"}
       }

    }

If ($ShowResults -like "InfoOnly")  {
    Write-Host "$BlockText $TestDate " -ForeGroundColor Yellow -NoNewLine
    Write-Host "$CheckName $TestCount" -ForegroundColor White
    $script:MailBody += "$BlockText $TestDate $CheckName $TestCount`r`n"
    }
}


Function UT-Get-PhoneSystemInfo {
<#
	.SYNOPSIS
	Get Teams Phone-system info of a RA, CQ or an  AA
    .DESCRIPTION
	Get Teams Phone-system info of ResourceAccount, DisplayName, CallQueue or AutoAttendant, Agent or phonenumber
	.COMPONENT
	UT-SfB
    .INPUTS
    DisplayLevel defines how specified the output is
    .OUTPUTS
    Outputs to display and can return a value
    Has also UTv1, UTv2, UTv3 as global variables
    .EXAMPLE
    UT-Get-PhoneSystemInfo -InputType DisplayName gogh
    .EXAMPLE
    UT-Get-PhoneSystemInfo -InputType UserName "v.gogh@utwente.nl"
    .EXAMPLE
    UT-Get-PhoneSystemInfo -InputType PhoneNumber "1234"
    .EXAMPLE
    PS> $CQtext =UT-Get-PhoneSystemInfo -InputType CQ -InputName  "All"
    PS> $CQtext | select Name,Identity, RoutingMethod, ApplicationInstances, Agents, OverflowAction, AllowOptOut, WelcomeMusicFileName, TimeoutAction, TimeoutThreshold, TimeoutActionTarget, PresenceBasedRouting |Export-Csv -NoTypeInformation C:\destination\CsCallQueue.csv
    .EXAMPLE
    PS> $AAtext =UT-Get-PhoneSystemInfo -InputType AA -InputName  "All"
    PS> $AAtext| select Name,Identity, LanguageId, VoiceId, DefaultCallFlow, Operator, TimeZoneId, CallFlows, Schedules, CallHandlingAssociations, ApplicationInstances |Export-Csv -NoTypeInformation C:\destination\CsAutoAttendant.csv
    .EXAMPLE
    PS> $RAtext = UT-Get-PhoneSystemInfo -InputType RA -InputName  "All"
    PS> $RAtext|Sort-Object UserPrincipalName| select DisplayName, PhoneNumber, UserPrincipalName, ObjectId |Out-File C:\destination\RA.csv

#>
	[CmdletBinding()]
	Param(
		[parameter(Position=0, Mandatory=$True, ValueFromPipeline=$False, HelpMessage='Give input name like +31534892222')]
			[String]$InputName,
		[parameter(Position=1, Mandatory=$False, ValueFromPipeline=$False, HelpMessage='Give input type')] [ValidateSet("PhoneNumber", "DisplayName", "UserName", "CQ", "RA", "AA")]
			[String]$InputType="",
		[parameter(Position=3, Mandatory=$False, ValueFromPipeline=$False, HelpMessage='Show extra info level')] [ValidateRange(0,9)] 
			[int]$DisplayLevel = 4
	)
#DisplayLevel  0  return essential output only 
#DisplayLevel  2  show standard output
#DisplayLevel >5  show extra agent info
#DisplayLevel Odd Also show comment messages

if (![bool]$InputType ) {
    switch -Regex ($InputName) {
	    '^(\+315348\d{5}$)|(^\d{4}$)'	{$InputType = "PhoneNumber"}
	     '@utwente.nl$' 				{$InputType = "UserName"} 
         '^CQ-' 						{$InputType = "CQ"}
         '^RA-' 						{$InputType = "RA"}
         '^AA-' 						{$InputType = "AA"}
        }
	if ($DisplayLevel%2) {UT_Display-Results -BlockText "Info0" -TestColor "White" -CheckName "InputType was empty, determined InputType: $InputType" }
	}

if ( ($InputType -eq "")  ) {  #-and ($InputName -eq "All")
    if ($DisplayLevel -ge 1) { Write-host "Please apply input type" -ForegroundColor Red }
    return 
    }
if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info0" -TestColor "White" -CheckName  "inputs are: Name: $InputName, type: $InputType + DisplayLevel: $DisplayLevel" }


if( -not (Get-Module -Name MicrosoftTeams)){
    Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
    }

if ( -not (Get-Module -Name SkypeOnlineConnector)) {
	Import-Module -Name SkypeOnlineConnector -ErrorAction SilentlyContinue
    }

if ( -not (Get-Module -Name MSOnline)) {
	Import-Module -Name MSOnline -ErrorAction SilentlyContinue
    }
# $SfBOnlineSession = New-CsOnlineSession -Credential $credentials -OverrideAdminDomain universiteittwente.onmicrosoft.com
# Import-PSSession $SfBOnlineSession -AllowClobber

# ( Get-AzureADTenantDetail|select -ExpandProperty VerifiedDomains|Where {$_.Capabilities -match  "Communication" }).Name
$domain = "@"+( Get-MsolDomain|Where {$_.Capabilities -match  "Communication" }).Name



$AllCQs = get-cscallqueue  -WarningAction silentlycontinue -First 250
$AllAAs = Get-CsAutoAttendant 
$AllRAs = Get-CsOnlineApplicationInstance

if (($InputType -eq "CQ") -and ($InputName -eq "All")) {Return ( $AllCQs ) }
if (($InputType -eq "AA") -and ($InputName -eq "All")) {Return ( $AllAAs ) }
if (($InputType -eq "RA") -and ($InputName -eq "All")) {Return ( $AllRAs ) }

If (($InputType -eq "CQ") -and( $InputName -ne "All")) {
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info1" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching CallQueue: $InputName" }
    # queueName  like CQ-MC-SI 
    $CQsFound = $AllCQs |Where-Object {$_.Name -match $InputName}
    foreach ($queue in $CQsFound) {
        $AppInst = $queue.ApplicationInstances[0]
        $RAsFound= $AllRAs|Where-Object {$_.ObjectId -match $AppInst }
        # Write-Host " `"$($RAsFound.DisplayName)`", `"$($RAsFound.PhoneNumber)`",  `"$($queue.Name)`" "
        #if ([bool]$RAcq) { $CQsFound[( $CQsFound.IndexOf($queue) )] | Add-Member -NotePropertyName ResAcc -NotePropertyValue $RAcq  -Force }
        }
    } # CQs
If (($InputType -eq "AA") -and( $InputName -ne "All")) {
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info1" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching AutoAttendant: $InputName" }
    $AAsFound = $AllAAs |Where-Object {$_.Name -match $InputName}
    foreach ($aa in $AAsFound) {
        $AppInst = $aa.ApplicationInstances[0]
        $RAsFound= $AllRAs|Where-Object {$_.ObjectId -match $AppInst }
        }
    } # AAs
     
If (($InputType -eq "RA") -and( $InputName -ne "All")) {
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info1" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching ResourceAccount: $InputName" }
    $RAsFound = $AllRAs |Where-Object {$_.UserPrincipalName -match $InputName}
    } # RAs
     
If ($InputType -eq "DisplayName") {
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info1" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching ResourceAccount with DisplayName: $InputName" }
    $RAsFound = $AllRAs |Where-Object {$_.DisplayName -match $InputName}
    } # DisplayName -> RAs
     
If ($InputType -eq "PhoneNumber") {
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info1" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching ResourceAccount with Phonenumber: $InputName" }
    $RAsFound = $AllRAs |Where-Object {$_.PhoneNumber -match $InputName}
    # put lines below in display results?
    if (![bool]$RAsFound) {
            if ($DisplayLevel -ge 1) {Write-Output "$InputType not found" -ForegroundColor Red}
            Return
            } else {
            if ($DisplayLevel%2) { Write-Output ("$InputType found in {0} queues " -f $RAsFound.count) }
            }

    } # PhoneNumber -> RAs


If ($InputType -eq "UserName") {
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info1" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching Agent with mailaddress: $InputName" }
     Try { 
        $agentGuid = (Get-CsOnlineUser $InputName).ObjectId.Guid # Try, error:Management object not found for identity
        }
     Catch { #catch werkt nog niet!!
           if ($DisplayLevel -ge 1) {Write-Output "$InputName is not an OnlineUser" -ForegroundColor Red}

        Return 
        }
     $CQsFound  =  @()
     ForEach ($cq in $AllCQs ) {
        if ($cq.agents.ObjectId -eq $agentGuid ) {
            if ($DisplayLevel%2) {UT_Display-Results -BlockText "Info2" -TestColor "White" -ShowResults "InfoOnly" -CheckName "$($cq.name)" }
            $CQsFound += ,$cq
            }
        }
        if ($CQsFound.count -eq 0) {
            if ($DisplayLevel -ge 1) {UT_Display-Results -BlockText "Info3" -TestColor "Red" -ShowResults "InfoOnly" -CheckName ""User not found in any CallQueues"" }
            Return 
            } 
    } # UserName -> CQ

# done searching primary input
 
if ( ($InputType -eq "RA") -or( $InputType -eq "DisplayName") -or ( $InputType -eq "PhoneNumber") ) { # if searching for RA then find belonging CQs and AAs
    $CQsFound = @()
    $AAsFound = @()
    $AutoAttendantID = "ce933385-9390-45d1-9512-c8d228074e07" 
    $CallQueueID     = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
    foreach ($RAFound in $RAsFound) {
        If ( ( [bool]$RAsFound ) -and ($RAsFound.ApplicationId -eq $CallQueueID)     ) { #  RA is for a CallQueue
            # find the matching CQ
            $CQsFound += $AllCQs |Where-Object {$_.ApplicationInstances -match $RAFound.ObjectId}
            } 
        If ( ( [bool]$RAsFound ) -and ($RAsFound.ApplicationId -eq $AutoAttendantID) ) { #  RA is for an AutoAttendant
            # find the matching AA
            $AAsFound += $AllAAs |Where-Object {$_.ApplicationInstances -match $RAFound.ObjectId}        
            }
        }
    } # RA -> CQs,AAs



# Add  info to each CallQueue
if (($DisplayLevel -ge 5) -and ([bool]$CQsFound)) { # or shouls I always add this, takes more time
    if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info-A" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching for Agents" }
    $agents=@()
    foreach ($queue in $CQsFound) {
        $QueueName = ($queue.Name)
        foreach ($agent in $($queue.agents) ) {
            $user  = $agent.ObjectId | Get-CsOnlineUser
            $OptIn = $queue.agents[( $agents.IndexOf($agent) )].OptIn
            $agent | Add-Member -NotePropertyName Name  -NotePropertyValue $user.UserPrincipalName -Force
            $agent | Add-Member -NotePropertyName First -NotePropertyValue $user.FirstName -Force
            $agent | Add-Member -NotePropertyName Last  -NotePropertyValue $user.LastName -Force
            $agent | Add-Member -NotePropertyName Uri   -NotePropertyValue $user.LineURI -Force
            $agent | Add-Member -NotePropertyName Department -NotePropertyValue $user.Department -Force
            $agent | Add-Member -NotePropertyName OptIn -NotePropertyValue $OptIn  -Force
            } 
        $TableProps = @{Property="Name","First","Last","Uri","Department","OptIn"}
        UT_Display-Results -BlockText "Queue" -CheckName "Agents in $QueueName" -TestColor "Green" -TestDisplay $Agents -ShowResults $ShowRes -TestTable $TableProps
        #$CQsFound[( $agents.IndexOf($agent) )] | Add-Member -NotePropertyName AgentsFull -NotePropertyValue $agents  -Force
        }
    }

 

if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info4" -TestColor "White" -ShowResults "InfoOnly" -CheckName "Searching Showing reults" }




# display result(s)
     if ($DisplayLevel%2)     {$ShowRes = "DataAndInfo" } # 3,5,7
     if ($DisplayLevel -eq 1) {$ShowRes = "InfoOnly"    }
     if ($DisplayLevel -ge 2) {$ShowRes = "DataOnly"    }
     if ($DisplayLevel -ge 8) {$ShowRes = "All"         }


if ( [bool]$CQsFound ) { # Output CQ
     if ($DisplayLevel -ge 2) {$TableProps = @{Property="Name"}                            }
     if ($DisplayLevel -ge 4) {$TableProps = @{Property="Name","RoutingMethod"}            }
     if ($DisplayLevel -ge 6) {$TableProps = @{Property="Name","Identity","RoutingMethod"} }
     if ($DisplayLevel -ge 8) {$TableProps = @{Property="Name","RoutingMethod","TimeoutAction","OverflowAction","AllowOptOut" } }
     if ($DisplayLevel -ge 9) {$TableProps = @() }
     UT_Display-Results -BlockText $InputType -CheckName "CallQueues" -TestColor "Green" -TestDisplay $CQsFound -ShowResults $ShowRes -TestTable $TableProps # -MaxResults $MaxResults -MaxWith $MaxWith

     <#if ($DisplayLevel -in (5..8) {
        foreach ($queue in $CQsFound) {
            UT_Display-Results -BlockText "Queue" -CheckName "Agents in $QueueName" -TestColor "Green" -TestDisplay $queue.AgentsFull -ShowResults $ShowRes -TestTable $TableProps
            }
      #>

      If ( ($CQsFound|Measure).Count -eq 1 ) { # display ResourceAccount belonging to CQ
        $AppInst = $CQsFound.ApplicationInstances[0]
        $CQResourceAccount  = $AllRAs |Where-Object {$_.ObjectId -match $AppInst}
        $TableProps = @{Property="DisplayName","PhoneNumber"} 
         UT_Display-Results -BlockText "RA_CQ" -CheckName "$($CQsFound.Name)" -TestColor "Green" -TestDisplay $CQResourceAccount -ShowResults $ShowRes -TestTable $TableProps # -MaxResults $MaxResults -MaxWith $MaxWith
        If (![bool]$CQResourceAccount.PhoneNumber) {
                if ($DisplayLevel%2) { UT_Display-Results -BlockText "Info5" -TestColor "White" -ShowResults "InfoOnly" -CheckName "$($CQResourceAccount.DisplayName) is probbably part of an AutoAttendant" }
            }
        }
    } # endif output CQ 

if ( [bool]$AAsFound )  { # Output AAs
     if ($DisplayLevel -ge 2) {$TableProps = @{Property="Name"}                            }
     if ($DisplayLevel -ge 4) {$TableProps = @{Property="Name","LanguageId"}               }
     if ($DisplayLevel -ge 6) {$TableProps = @{Property="Name","Identity","LanguageId"}    }
     if ($DisplayLevel -ge 8) {$TableProps = @{Property="Name","Identity","LanguageId","Schedules","CallFlows" } }
     if ($DisplayLevel -ge 9) {$TableProps = @()                                           }
     UT_Display-Results -BlockText $InputType -CheckName "AutoAttendants" -TestColor "Green" -TestDisplay $AAsFound -ShowResults $ShowRes -TestTable $TableProps # -MaxResults $MaxResults -MaxWith $MaxWith
    #$ThisAA.Schedules)[0].WeeklyRecurrentSchedule
    }

if ([bool]$RAsFound ) { # Output ResourceAccounts
     if ($DisplayLevel -ge 2) {$TableProps = @{Property="DisplayName"}                          }
     if ($DisplayLevel -ge 4) {$TableProps = @{Property="DisplayName","PhoneNumber"}            }
     if ($DisplayLevel -ge 6) {$TableProps = @{Property="DisplayName","PhoneNumber","ObjectId"} }
     if ($DisplayLevel -ge 9) {$TableProps = @()                                                }
     UT_Display-Results -BlockText "ResAcc" -CheckName "ResourceAccount" -TestColor "Green" -TestDisplay $RAsFound -ShowResults $ShowRes -TestTable $TableProps # -MaxResults $MaxResults -MaxWith $MaxWith
}

if ($DisplayLevel -ge 1) { Write-Host "Examine variables `$UTv1, `$UTv2 and `$UTv3 for more details.`r`n" }
$Global:UTv1 = $RAsFound
$Global:UTv2 = $CQsFound
$Global:UTv3 = $AAsFound

# Remove-PSSession $SfBOnlineSession
# Disconnect-MicrosoftTeams


} #end function UT-Get-PhoneSystemInfo


Function UT-New-CallQueue {
<#
	.SYNOPSIS
	Create a new Teams Phone-system CallQueue
	.DESCRIPTION
	Create a new Teams Phone-system CallQueue
	.COMPONENT
	UT-SfB
    .EXAMPLE
    PS> UT-New-CallQueue -CQname "CQ-LISA-test-4" -DisplayName  "Lisa CQ test 4" -PhoneNumber "+31534891123"
#>
	[CmdletBinding()]
	Param(
		[parameter(Position=0, Mandatory=$True, ValueFromPipeline=$False, HelpMessage='Give Name like CQ-TNW-Secretariat-AST, CQ-CFM-SD-ITC')] #Faculty or service departments - groupname - purpose 
			[String]$CQname,
		[parameter(Position=1, Mandatory=$False, ValueFromPipeline=$False, HelpMessage='Give Displayname like creditors administration')] # list, table, xml, csv
			[String]$DisplayName,
		[parameter(Position=2, Mandatory=$False, ValueFromPipeline=$False, HelpMessage='like +31534891234')]  
			[String]$PhoneNumber="" 
	)
#PhoneNumber ValidateSet '^(\+315348\d{5}$)

#$InputName       = "SV-Atlantis"
#$DisplayName     = "S.V. Atlantis"
#$PhoneNumber             = "+31534891234"

$AutoAttendantID = "ce933385-9390-45d1-9512-c8d228074e07" 
$CallQueueID     = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
$RA              = "RA-"+$CQname+$Domain

If ($CQname -notlike "CQ-*") { Return "Name should be like CQ-"+$CQname }
If ([bool]$RaPhoneNr) {
    If ($RaPhoneNr -like "Tel:*") { Return "Try without Tel:" }
    If ($RaPhoneNr -notlike "+*") { Return "Name Phonenumber be like +31534891234" }
    }

if( -not (Get-Module -Name MicrosoftTeams)){
    Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
    }
if ( -not (Get-Module -Name SkypeOnlineConnector)) {
	Import-Module -Name SkypeOnlineConnector -ErrorAction SilentlyContinue
    }
if ( -not (Get-Module -Name MSOnline)) {
	Import-Module -Name MSOnline -ErrorAction SilentlyContinue
    }


if ([bool]$InputName) {
    New-CsOnlineApplicationInstance -UserPrincipalName $RA -ApplicationId $CallQueueID -DisplayName $DisplayName
    New-CsCallQueue -Name $CQname -RoutingMethod Attendant -OverflowThreshold 5 -UseDefaultMusicOnHold $true -LanguageId "en-GB"

    write-host "Sleeping 30 sec."
    Start-Sleep -s 30
    # ResourceAccount, add location and license
    Set-MsolUser -UserPrincipalName $RA  -UsageLocation NL
    Set-MsolUserLicense -UserPrincipalName $RA  -AddLicenses "universiteittwente:PHONESYSTEM_VIRTUALUSER_FACULTY"

    # Associate ResourceAccount with  CallQueue / autoAttendantId 
    write-host "Sleeping 30 sec."
    Start-Sleep -s 30
    $ObjectId = (Get-CsOnlineUser $RA).ObjectId    # 1d90c2dc-b6b0-4f80-93c8-d6565cb33782
    $QueueId               = (Get-CsCallQueue -NameFilter $CQname).Identity                 # ac7529c1-6418-442b-a780-fd90b6a3a250
    New-CsOnlineApplicationInstanceAssociation -Identities @($ObjectId) -ConfigurationId $QueueId         -ConfigurationType CallQueue    
    }

if ([bool]$PhoneNumber) { # set phonenumber for ResourceAccount
    #Start-Sleep -s 40
    Set-CsOnlineApplicationInstance -Identity $RA  -OnpremPhoneNumber $PhoneNumber
    }

write-host "Now go to the Microsoft Teams admin center add some agents "

} #  end function UT-New-CallQueue



Function UT-New-EmptyCsAutoAttendant {
<#
	.SYNOPSIS
	Create new AutoAttendant for Teams 
    .DESCRIPTION
	Create new Empty AutoAttendant for Teams Phone-system
	.COMPONENT
	UT-SfB
    .EXAMPLE
    PS> UT-New-EmptyCsAutoAttendant -AaName "AA-Lisa-test-4" -RaDisplayName "Lisa AA test 4" -PhoneNumber "+31534891123"
#>
[CmdletBinding()]
Param(
	[parameter(Position=0, Mandatory=$True, ValueFromPipeline=$False, HelpMessage='Give AutoAttendant name like AA-LISA-Test2')]
		[String]$AaName,
	[parameter(Position=1, Mandatory=$False, ValueFromPipeline=$False, HelpMessage='Give DisplayName')] 
		[String]$RaDisplayName,
	[parameter(Position=3, Mandatory=$False, ValueFromPipeline=$False, HelpMessage='Give Phonenumber like +31534891234')]  
		[String]$PhoneNumber 
)
#PhoneNumber ValidateSet '^(\+315348\d{5}$)


If ($AaName -notlike "AA-*") { Return "Name should be like AA-"+$AAName }
If ([bool]$PhoneNumber) {
    If ($PhoneNumber -like "Tel:*") { Return "Try without Tel:" }
    If ($PhoneNumber -notlike "+*") { Return "Name Phonenumber be like +31534891234" }
    }
# $SfBOnlineSession = New-CsOnlineSession -Credential $credentials -OverrideAdminDomain universiteittwente.onmicrosoft.com
# Import-PSSession $SfBOnlineSession -AllowClobber
if( -not (Get-Module -Name MicrosoftTeams)){
    Import-Module MicrosoftTeams -ErrorAction SilentlyContinue
    }
if ( -not (Get-Module -Name SkypeOnlineConnector)) {
	Import-Module -Name SkypeOnlineConnector -ErrorAction SilentlyContinue
    }
if ( -not (Get-Module -Name MSOnline)) {
	Import-Module -Name MSOnline -ErrorAction SilentlyContinue
    }

$domain = "@"+( Get-MsolDomain|Where {$_.Capabilities -match  "Communication" }).Name

$RaAaName = "RA-"+$AAName
$RaAaUPN  = $RaAaName+$domain
#$RaCqName = "RA-CQ-"+$AAName.Substring(6)
$AutoAttendantID = "ce933385-9390-45d1-9512-c8d228074e07" 
$CallQueueID     = "11cd3e2e-fccb-42ad-ad00-878b93575e07"


If ([bool]$PhoneNumber) { 
    Try {$FoundRaPhoneNr = (Get-CsOnlineApplicationInstance -ErrorAction Silentlycontinue |? {$_.PhoneNumber -like "*$PhoneNumber"} ) }
    Catch {write-host "$PhoneNumber does not exists"}
    If ( ( [bool]$FoundRaPhoneNr ) -and ($FoundRaPhoneNr.ApplicationId -eq $AutoAttendantID) ) { # found an AutoAttendant RA, 
        Write-Host "Resourceaccount $($FoundRaPhoneNr.UserPrincipalName) for $PhoneNumber already exists" 
        Return ""
        }
    If ( ( [bool]$FoundRaPhoneNr ) -and ($FoundRaPhoneNr.ApplicationId -eq $CallQueueID)     ) { # found a CallQueue RA, 
        $CQ = get-cscallqueue -WarningAction silentlycontinue -First 250 |Where-Object {$_.ApplicationInstances -match $FoundRaPhoneNr.ObjectId} # find the matching CQ
        Write-Host "Phonenumber $PhoneNumber belongs to Callqueue $($CQ.DisplayName)"
        # clear phonenumber for Callqueue
        Set-CsOnlineApplicationInstance -Identity $($FoundRaPhoneNr.ObjectId)  -OnpremPhoneNumber $null
        $CQident = $FoundRaPhoneNr.ObjectId # if Callqueue was found with the phonenumber
        } 
    } #phonenumber does not exist (anymore)
        
Try
    {
    Get-CsOnlineUser "$RaAaName" -ErrorAction Silentlycontinue #if (!($ErrorMessage -match "Management object not found")) {Write-Host "Skipping this CallQueue because the AA already exists."} 
    }
Catch # the AA does not already exists
    {

    # create ResourceAccount for AA
    New-CsOnlineApplicationInstance -UserPrincipalName "$RaAaUPN" -ApplicationId $AutoAttendantID -DisplayName $RaDisplayName
    Write-Host "Sleeping for 30sec."
    Start-Sleep 30 
    Set-MsolUser -UserPrincipalName $RAaaUPN  -UsageLocation NL
    Set-MsolUserLicense -UserPrincipalName "$RAaaUPN"  -AddLicenses "universiteittwente:PHONESYSTEM_VIRTUALUSER_FACULTY"

    # Create AutoAttendant
        # $greetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Welcome to the  Servicedesk"
        # $RAcqObectId = ...xxx...
        # $CQappinstance = (Get-CsOnlineUser -Identity "CQident").ObjectId # one of the application instances associated to the Call Queue
    if ([bool]$CQident) { 
        $CallQ1 = New-CsAutoAttendantCallableEntity -Type ApplicationEndpoint -Identity $CQident 
        $menuOptionZero = New-CsAutoAttendantMenuOption -DtmfResponse Automatic -Action TransferCallToTarget -CallTarget  $CallQ1 # -DtmfResponse Tone0
        } else
        {
        $menuOptionZero = New-CsAutoAttendantMenuOption -DtmfResponse Automatic -Action DisconnectCall -CallTarget  $CallQ1 # -DtmfResponse Tone0
        }
        # $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Hi"
    $defaultMenu = New-CsAutoAttendantMenu -Name "Default menu"  -MenuOptions @($menuOptionZero) #   -Prompts @($menuPrompt) -EnableDialByName  -DirectorySearchMethod 
    $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default call flow" -Menu $defaultMenu # -Greetings @($greetingPrompt)
    New-CsAutoAttendant -Name $AAname  -Language "en-GB" -TimeZoneId "W. Europe Standard Time" -Operator $null  -DefaultCallFlow $defaultCallFlow #  -EnableVoiceResponse  -CallFlows @($afterHoursCallFlow) -CallHandlingAssociations @($afterHoursCallHandlingAssociation)  -InclusionScope $inclusionScope

    # Associate ResourceAccount with  CallQueue / autoAttendantId 
    $RAAAObjectId = (Get-CsOnlineUser "$RAaaName").ObjectId   
    $AutoAttendantId        =(Get-CsAutoAttendant -NameFilter $AAname).Identity               
    New-CsOnlineApplicationInstanceAssociation -Identities @($RAAAObjectId) -ConfigurationId $AutoAttendantId -ConfigurationType AutoAttendant

    # set phonenumber for ResourceAccount
    } # end catch

    If ([bool]$PhoneNumber) { Set-CsOnlineApplicationInstance -Identity $RAaaUPN  -OnpremPhoneNumber $PhoneNumber }
    

} #  end function New-EmptyCsAutoAttendant

