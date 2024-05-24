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
    #Write-output $Inputs

    #matches variables to inputs from the webhook -- THANK YOU DANIEL 
    $Firstname = $inputs.content.FirstName
    $Lastname = $inputs.content.LastName
    $Department = $inputs.content.Department
    $Title = $inputs.content.Title
    $Supervisor = $inputs.content.Supervisor
    #seperates the supervisor name into two parts
    $SupFirstname, $SupLastname = $Supervisor.split()
    #sets the upn for the supervisor
    $Manager = $SupFirstname.SubString(0,1)+$SupLastname+'@linq.com'
    #substring collects the first letter of $Firstname
    $UPN = $Firstname.SubString(0,1)+$Lastname+'@linq.com'
    #makes the $UPN variable all lowercase
    $UPNlower = $UPN.ToLower()
    $Displayname = $Firstname+' '+$Lastname
    $EmploymentType = $inputs.content.EmploymentType
    $MailNickName = $Firstname.SubString(0,1)+$Lastname
    $MailNickNamelower = $MailNickName.ToLower()
    #write-output $MailNickNamelower

}
else {
    Write-Output 'No data received...Exiting'
    exit
}


#variables had to be put down here AFTER the parameter field
$ClientID = Get-AutomationVariable -Name 'ClientID'
$TenantId = Get-AutomationVariable -Name 'TenantID'
$Thumbprint = Get-AutomationVariable -Name 'Thumbprint'
$PasswordProfileEmployee = Get-AutomationVariable -Name 'PasswordProfileEmployee'
$PasswordProfileContractor = Get-AutomationVariable -Name 'PasswordProfileContractor'
$EMSE3 = Get-AutomationVariable -Name 'EMSE3License'
$Office365 = Get-AutomationVariable -Name 'Office365License'
$ADP1 = Get-AutomationVariable -Name 'P1license'
$M365LicenseApply = Get-AutomationVariable -Name 'M365License'
$APPID = Get-AutomationVariable -Name 'EXOnlineAPPID'
$FullTimeGroups = @('54464eab-dc68-45d0-af1a-e6719cabdf62','28c379aa-9df0-40c5-90c9-08e0c1f8140a','87f3b8b0-0b91-40d3-ab31-f15de498d6a7')
$ContractorGroups = @('54464eab-dc68-45d0-af1a-e6719cabdf62','6092c807-95e8-4d21-bace-71099df9fa06')
$FullTimeMailEnabledGroups = @('f2153a76-48ce-4a6b-a490-0a5e858dfb4e')
#combines both licenses into one variable
$EmployeeLicenses = @(
    @{SkuId = $EMSE3}
    @{SkuId = $Office365}
)
$ContractorLicenses = @(
    @{SkuId = $ADP1}
    @{SkuId = $Office365}

)


$EmployeePassword = @"
{
    "forceChangePasswordNextSignIn": false,
    "forceChangePasswordNextSignInWithMfa": false,
    "password": "$PasswordProfileEmployee"
}
"@


$ContractorPassword = @"
{
    "forceChangePasswordNextSignIn": false,
    "forceChangePasswordNextSignInWithMfa": false,
    "password": "$PasswordProfileContractor"
}
"@

function GroupAdditions{
    try {
        foreach ($GroupObjectID in $FullTimeGroups)
        {New-MgGroupMember -GroupId $GroupObjectID -DirectoryObjectID UserIDHere}
            
    }
    catch {
        <#Do this if a terminating exception happens#>
    }
}


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
        "activityTitle": "$Displayname -- Provision Update",
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





#connects to MS graph powershell
Connect-MgGraph -ClientId $ClientID -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

#checks if the user already exists
Write-output "Checking for existing user..."
$UserCheck = (get-mguser -userid $UPNlower -ErrorAction 'SilentlyContinue').UserPrincipalName 

#user exists, leave script
if ($Usercheck -eq $UPNlower) {
    Write-output "$UPNlower already exists. Exiting script...."        ######THIS BLOCK CAN BE IMPROVED.. ADDING EXTRA STEPS TO CHANGE THE UPN FOR FirstinitialSecondinitialLastname@linq.com instead of quitting the script
    write-output "Disconnecting..."
    Disconnect-MGgraph
    #$Notes = "$UPNLower already exists."
    TeamsNotif -Status 'Failed' -Notes "$UPNLower already exists"
}
else {
    #user does not exist, continue.
   write-output "User does not exist, running script."
}


#Checks if the employment type field is selected for full time or contractor/intern. If Employee, run this, Else run contractor commands
If(($EmploymentType -eq "Employee (full time)") -or ($EmploymentType -eq "Intern") -or ($EmploymentType -eq "Employee (part time)"))
{
   
    Write-Output "User is marked as $EmploymentType"
    #collects the id for the new user
    #$UserID = (Get-MgUser -userid $UPNlower).id           ---not needed
    #collects the ID for the Manager
    $ManagerID = (Get-Mguser -property id -UserId $Manager).id
    #Uses the Id and attached it to a string
    $ApplyManager = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$ManagerID"}
    #Creates the user & applies password config
    #try statement to catch errors when making the user
    try{
        #creates the user
        New-MgUser -DisplayName $Displayname -PasswordProfile $EmployeePassword -AccountEnabled -MailNickName $MailNickNamelower -UserPrincipalName $UPNlower -UsageLocation US -Department $Department -JobTitle $Title -CompanyName "EMS LINQ, Inc." -GivenName $Firstname -Surname $Lastname -ErrorAction 'Stop'
        #Adds user to requested groups, not fleshed out yet... -- should add variables with groupid's contianed, like the licenses
        #New-MgGroupMember -GroupId GroupIDHere -DirectoryObjectID UserIDHere
    }
    catch{
        Write-Output "Error Encountered:"$_
        write-output "Disconnecting..."
        Disconnect-MGgraph
        #$Notes = "Error occured creating the user"
        TeamsNotif -Status 'Failed' -Notes "Error occured creating the user"
    }
    #sets the manager in AAD
    Set-mgUsermanagerbyref -UserId $UPNlower -BodyParameter $ApplyManager
    #sets basic licenses
    try {
        #this is commented out TEMPORARILY. We have 100 or so m365 licenses that need to be used. This coveres the 
        #Set-MgUserLicense -UserId $UPN -AddLicenses $EmployeeLicenses -RemoveLicenses @() -ErrorAction 'Stop'
        Set-MgUserLicense -UserId $UPN -AddLicenses $EmployeeLicenses -RemoveLicenses @() -ErrorAction 'Stop'
        Write-Output 'Created Full Time Employee and applied EMS E3 license'
    }
    catch {
        Write-Output "Error Encountered:"$_
        write-output "Disconnecting..."
        Disconnect-MGgraph
        #$Notes = "Error occured creating the user"
        TeamsNotif -Status 'Failed' -Notes "Could not apply one or more licenses"
    }

    $UserObjID =(Get-Mguser -property id -UserId $UPNlower).id
    foreach($GroupObjectID in $FullTimeGroups)
    {
        try {
            New-MgGroupMember -GroupId $GroupObjectID -DirectoryObjectID $UserObjID -ErrorAction 'Stop'
        }
        catch {
            Disconnect-MgGraph
            TeamsNotif -Status 'Failed' -Notes "Could not add to Group ID: $GroupObjectID"
        }
    }

    Disconnect-MGgraph

    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $APPID -Organization "emslinqinc.onmicrosoft.com"
    try {
        write-output "attempting to add to company group"
        Add-DistributionGroupMember -identity $FullTimeMailEnabledGroups -member $UPNlower -ErrorAction 'Stop'
    }
    catch {
        Disconnect-ExchangeOnline -Confirm:$false
        TeamsNotif -Status 'Failed' -Notes "Could not add to Company Group"
    }
    
    
} 
elseif($EmploymentType -eq "Contractor") #Contractor commands -- Differences: passwordprofile, licenses, groups.
{
    
    Write-Output "User is marked as $EmploymentType"
    #$UserID = (Get-MgUser -userid $UPNlower).id      --not needed
    $ManagerID = (Get-Mguser -property id -UserId $Manager).id
    $ApplyManager = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$ManagerID"}
    try{
        New-MgUser -DisplayName $Displayname -PasswordProfile $ContractorPassword -AccountEnabled -MailNickName $MailNickNamelower -UserPrincipalName $UPNlower -UsageLocation US -Department $Department -JobTitle $Title -CompanyName "EMS LINQ, Inc." -GivenName $Firstname -Surname $Lastname -ErrorAction 'Stop'
        #Addes user to requested groups, not fleshed out yet... -- should add variables with groupid's contianed, like the licenses
        #New-MgGroupMember -GroupId groupidhere -DirectoryObjectID UserIDHere
    }
    catch{
        write-output "Error Encountered:"$_
        write-output "Disconnecting..."
        Disconnect-MGgraph
        TeamsNotif -Status 'Failed' -Notes "Could not create $EmploymentType"
    }


    Set-mgUsermanagerbyref -UserId $UPNlower -BodyParameter $ApplyManager

    try {
        Set-MgUserLicense -UserId $UPN -AddLicenses $ContractorLicenses -RemoveLicenses @()
        Write-output 'Created Contractor and applied AD license'
    }
    catch {
        write-output "Error Encountered:"$_
        write-output "Disconnecting..."
        Disconnect-MGgraph
        TeamsNotif -Status 'Failed' -Notes "Could not create $EmploymentType"
    }
    
    $UserObjID =(Get-Mguser -property id -UserId $UPNlower).id
    foreach($GroupObjectID in $ContractorGroups)
    {
        try {
            New-MgGroupMember -GroupId $GroupObjectID -DirectoryObjectID $UserObjID -ErrorAction 'Stop'
        }
        catch {
            Disconnect-MgGraph
            TeamsNotif -Status 'Failed' -Notes "Could not add to $GroupObjectID"
        }
    }

   
    
}
else  #if neither, exit program.
{
   
    Write-Output "Employment Type not recognized"

}




write-output "Disconnecting..."
Disconnect-MGgraph
Disconnect-ExchangeOnline -Confirm:$false
TeamsNotif -Status 'Success!'
