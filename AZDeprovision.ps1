param (
    [Parameter (Mandatory = $false)]
    [object] $WebHookData
)

if ($WebHookData){
    # Header message passed as a hashtable 
    Write-Output "The Webhook Header Message"
    Write-Output $WebHookData.RequestHeader.Message

    # This is the name of the webhook when configured in Azure Automation
    Write-Output 'The Webhook Name'
    Write-Output $WebHookData.WebhookName

    # Body of the message.
    Write-Output 'The Request Body'
    Write-Output $WebHookData.RequestBody

    #Custom code
    #transfers webhook data into whatever this is?
    $Inputs = ConvertFrom-Json $WebHookData.Requestbody

    #$ADSupervisor =$inputs.content.Supervisor
    #$Title = $Inputs.content.Title
    $Employee = $Inputs.content.Employee
    $UPN = $Inputs.content.ExactEmail
    $UPNLower = $UPN.ToLower()
    $EmailForward =$Inputs.content.ForwardEmailAddress
    $CalendarRemoval = $Inputs.content.CalendarEvents
    

}
else {
    Write-Output 'No data received...Exiting'
    exit
}
#information to connect to MSGraph
$ClientID = Get-AutomationVariable -Name 'ClientID'
$TenantId = Get-AutomationVariable -Name 'TenantID'
$Thumbprint = Get-AutomationVariable -Name 'Thumbprint'
$APPID = Get-AutomationVariable -Name 'EXOnlineAPPID'
$ExchangeLic = Get-AutomationVariable -Name 'ExchangeLicense'
#parameter prep for disabling the account
$AccDisable = @{ AccountEnabled = "false"}
#sets up an empty list to remove distribution lists
$DLCheck = New-object System.Collections.Generic.List[System.Object]


#function to send off the teams notification with json formatting. 
function TeamsNotif{
    
   param (
        [string]$Status,
        [string]$Notes
    )
  
    $TeamsINCWebhook = Get-AutomationVariable -Name 'TeamsWebhook'
    $Date = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'Eastern Standard Time')
    
    $jsonParams = @"
    {
    "@type": "MessageCard",
    "themeColor": "0076D7",
    "summary": "Provision/Deprovision Updates",
    "sections": [{
        "activityTitle": "$Employee -- Deprovision Update",
        "activitySubtitle": "Via Azure Automate",
        "activityImage": "https://linqwpdevstor.blob.core.windows.net/desktopwallpaper/_LINQ_Icon_FullColor_AlphaBG_RGB.png",
        "facts": [{
            "name": "UPN:",
            "value": "$UPNLower"
        }, {
            "name": "Time Completed:",
            "value": "$Date EST"
        }, {
            "name": "Status:",
            "value": "$Status"
        }, {
            "name": "Errors:",
            "value": "$_"
        }, {
            "name": "Notes:",
            "value": "$Notes"
        }],
        "markdown": true
        }],

}
"@

    Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $jsonParams -Uri $TeamsINCWebhook
    write-Output "Exiting..."
    exit
}

function LicenseAssignment {

    foreach ($License in $ActiveLicenses)
    {
        write-output "running licenseAssignment function"
        write-output "Active license: $License"
        try {
            Set-MgUserLicense -userid $UPNLower -AddLicenses @() -RemoveLicenses @($License) -ErrorAction 'Stop'
        }
        catch {
            Write-Output "Error Encountered:"$_
            write-output "Disconnecting from exchange online..."
            Disconnect-MGgraph
            write-Output "Disconnected..."
            TeamsNotif -Status 'Failed' -Notes "Could not remove a license"
        }
        
        write-output "User removed from $License"
    }
    
}


#This function has to be done BEFORE licenses get removed, otherwise email forwarding will fail from a mailbox that does not exist.
function ApplyEmailForwarding {
    write-output "running email forwarding fucntion"
    write-output "connecting to exchange online"
    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $APPID -Organization "emslinqinc.onmicrosoft.com"
    write-output "applying email forwarding & hiding from GAL"
    set-mailbox -identity $UPNLower -ForwardingAddress $EmailForward -HiddenFromAddressListsEnabled $true
    write-output "disconnecting from exchange online"
    Disconnect-ExchangeOnline -Confirm:$false
    
}


#connects to MSGraph
try {
    Write-output "Connecting to MSGraph"
    Connect-MgGraph -ClientId $ClientID -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome
}
catch {
    Write-Output "Error Encountered:"$_
    write-output "Disconnecting..."
    Disconnect-MGgraph
    TeamsNotif -Status 'Failed' -Notes "Unable to connect to MSGraph"
}


#disables the user account
try {
    update-mguser -userid $UPNLower -BodyParameter $AccDisable -ErrorAction 'Stop'
}
catch {
    Write-Output "Error Encountered:"$_
    write-output "Disconnecting..."
    Disconnect-MGgraph
    TeamsNotif -Status 'Failed' -Notes "Could not disable the user"
}


#Sets up variables for directory ID's
$UserID = (Get-Mguser -userid $UPNLower).id
$GroupList = (Get-MguserMemberOf -userid $UPNLower).id
#sets a variable for each group ID

try {
    #removes manager
    write-output "Removing manager"
    Remove-MgUserManagerByRef -userid $UserID
}
catch {
    Write-Output "Error Encountered:"$_
    write-output "Unable to remove manager"
    Disconnect-MGgraph
    TeamsNotif -Status 'Failed' -Notes "Unable to remove the Manager"
}

foreach ($Group in $GroupList)
{
    try {
        #removes each group the user is apart of, via ID's
        #write-output "user is part of $Group"
        Remove-MgGroupMemberByRef -GroupId $Group -DirectoryObjectId $UserID -ErrorAction 'Stop'     
        write-output  "removed user from $Group"
    }
    catch{
        Write-Output "unable to remove from $Group. Added to DL List"
        #adds leftover groups into a list
        $DLCheck.add($Group)
    }
    
}



Write-output "DL List: $DLCheck"


#checks what licenses are applied to the user
$ActiveLicenses = (Get-MgUserLicenseDetail -userid $UPNLower).SkuId
#assigns a variable to each SkuID.

if ($null -eq $EmailForward) {
    write-output "running empty if statement"
    LicenseAssignment
}
elseif ($null -ne $EmailForward) {
    write-output "running full if statement"
    ApplyEmailForwarding
    try {
        #this is created after the license check so it isn't included when removed from the LicenseAssignment fucntion
        write-output "removing licenses and applying exchange online license"
        Set-MgUserLicense -userid $UPNLower -AddLicenses @{SkuId = $ExchangeLic} -RemoveLicenses @() -ErrorAction 'Stop'
    }
    catch {
        Write-Output "Error Encountered:"$_
        write-output "Disconnecting from exchange online..."
        Disconnect-MGgraph
        write-Output "Disconnected..."
        TeamsNotif -Status 'Failed' -Notes "Could not assign an Exchange License"
    }
    LicenseAssignment
}



#disconnecting from MGgraph
write-output "Disconnecting from MSGraph"
Disconnect-MGgraph



#if DLCheck isn't empty, run.
write-output "DLcheck $DLCheck"
# if ($null -ne $DLCheck) 
if ($DLCheck.Count -ge 1)
{
    
    try {
       #connects to exchange online
        write-output "Connecting to Exchange online"
        Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $APPID -Organization "emslinqinc.onmicrosoft.com"
        write-output "Connected"
    }
    catch {
        #exit if unable to connect to exchange online
        Write-Output "Error Encountered:"$_
        write-output "Disconnecting from exchange online..."
        Disconnect-ExchangeOnline -Confirm:$false
        write-Output "Disconnected..."
        TeamsNotif -Status 'Failed' -Notes "Could not connect to exchange online"
        
    }
    #for each group, remove the user. 
    foreach($DL in $DLCheck)
    {
        if ($DL -ne "f2153a76-48ce-4a6b-a490-0a5e858dfb4e") {
        write-output $DL
        #Removes DL's user is apart of
        Remove-DistributionGroupMember -identity $DL -member $UPNLower -Confirm:$false
        write-output "attempting to remove from DL's"
        }
        else {
            "Skipped Company group removal"
        }
        
    }
    
}
elseif($DLCheck.count -eq 0)
{
    write-output "There are no DL's to remove. Exiting..."
}

if ($CalendarRemoval -eq "Yes") {
    write-output "removing Calendar events"
    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $APPID -Organization "emslinqinc.onmicrosoft.com"
    try {
        Remove-CalendarEvents -identity $UPNLower -CancelOrganizedMeetings -QueryWindowInDays 90 -Confirm:$false
    }
    catch {
        Write-Output "Error Encountered:"$_
        write-output "Disconnecting from exchange online..."
        Disconnect-ExchangeOnline -Confirm:$false
        write-Output "Disconnected..."
        TeamsNotif -Status 'Failed' -Notes "Could not cancel meetings"
    }
    
}

Disconnect-ExchangeOnline -Confirm:$false
TeamsNotif -Status 'Success!'
