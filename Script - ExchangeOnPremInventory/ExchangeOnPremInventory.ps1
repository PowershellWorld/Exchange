param ([string]$Outputpath)
If ($Outputpath -eq "") {
    $Outputpath = $env:TEMP
}


Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "##############################################################################################################" -ForegroundColor Green
Write-Host ""
Write-Host "Exchange OnPremises inventory for domain $($ENV:USERDOMAIN)" -ForegroundColor Yellow -BackgroundColor Red
Write-Host "version 1.1.2 created by Alex ter Neuzen for PowershellWorld" -ForegroundColor White
Write-Host "   __________                                 .__           .__  .__  __      __            .__       .___   "  
Write-Host "   \______   \______  _  __ ___________  _____|  |__   ____ |  | |  |/  \    /  \___________|  |    __| _/   " 
Write-Host "   |     ___/  _ \ \/ \/ // __ \_  __ \/  ___/  |  \_/ __ \|  | |  |\   \/\/   /  _ \_  __ \  |   / __ |     " 
Write-Host "   |    |  (  <_> )     /\  ___/|  | \/\___ \|   Y  \  ___/|  |_|  |_\        (  <_> )  | \/  |__/ /_/ |     " 
Write-Host "   |____|   \____/ \/\_/  \___  >__|  /____  >___|  /\___  >____/____/\__/\  / \____/|__|  |____/\____ |     " 
Write-Host "                              \/           \/     \/     \/                \/                         \/     " 
#Write-Host "                         script is downloaded from https://www.powershellworld.com"
Write-host ""
Write-Host "Script start time: $(Get-Date)"
Write-Host ""



#$OUTPUTPATH = "C:\TEMP\OUTPUT"


#region Exchange Connection
Write-Host "Checking for Exchange Server in $($ENV:USERDOMAIN)"
Write-Host ""
Function Get-ADExchangeServer {
    # first a quick function to convert the server roles to a human readable form
    Function ConvertToExchangeRole {
        Param(
            [Parameter(Position = 0)]
            [int]$roles
        )
        $roleNumber = @{
            2  = 'MBX';
            4  = 'CAS';
            16 = 'UM';
            32 = 'HUB';
            64 = 'EDGE';
        }
        $roleList = New-Object -TypeName Collections.ArrayList
        foreach ($key in ($roleNumber).Keys) {
            if ($key -band $roles) {
                [void]$roleList.Add($roleNumber.$key)
            }
        }
        Write-Output $roleList
    }
  
    # Get the Configuration Context
    $rootDse = Get-ADRootDSE
    $cfgCtx = $rootDse.ConfigurationNamingContext
  
    # Query AD for Exchange Servers
    $exchServers = Get-ADObject -Filter "ObjectCategory -eq 'msExchExchangeServer'" `
        -SearchBase $cfgCtx `
        -Properties msExchCurrentServerRoles, networkAddress, serialNumber
    foreach ($server in $exchServers) {
        Try {
            $roles = ConvertToExchangeRole -roles $server.msExchCurrentServerRoles
  
            $fqdn = ($server.networkAddress | 
                Where-Object { $_ -like 'ncacn_ip_tcp:*' }).Split(':')[1]
  
            New-Object -TypeName PSObject -Property @{
                Name        = $server.Name;
                DnsHostName = $fqdn;
                Version     = $server.serialNumber[0];
                ServerRoles = $roles;
            }
        }
        Catch {
            Write-Error "ExchangeServer: [$($server.Name)]. $($_.Exception.Message)"
        }
    }
}

$ExchangeServer = (Get-ADExchangeServer).DNSHostName

If ($ExchangeServer -eq $null) {
    Write-host "- no exchange server was found" -BackgroundColor Red -ForegroundColor Yellow
    break
}
else {
    Write-host "- an Exchange server was found with $Exchangeserver as DNS. Trying to connect"
    $Sessions = Get-PSSession
    If ($($Sessions.ConfigurationName) -contains "Microsoft.Exchange") {
        Write-Host "- Exchangeserver is allready connected"
    }
    else {
        Try {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos -ErrorAction SilentlyContinue
            $connect = Import-PSSession $Session -DisableNameChecking -AllowClobber -ErrorAction SilentlyContinue
        
            if (!(Get-PSSession | where { $_.ConfigurationName -eq "Microsoft.Exchange" })) {
                Write-host "- Exchange environment is not connected" -ForegroundColor Red -BackgroundColor Yellow
                
            }
            else {
                Write-host "- Exchange environment is connected" -ForegroundColor green
            }
        }
        Catch {
            Write-Host "- There was an error connecting to $exchangeserver" -ForegroundColor Red -BackgroundColor Yellow
            break
        }
    }
}
#endregion Exchange Connection

Write-Host ""

$DomainControllerToUse = (Get-ADDomainController).HostName

Write-Host "Using $($DomainControllerToUse) as direct domain controller"

$TableUsers = @()
$RecordUsers = [ordered]@{
    "Name"                = ""
    "SamAccountName"      = ""
    "DisplayName"         = ""
    "UPN"                 = ""
    "PrimarySMTPAddress"  = ""
    "Aliasses"            = ""
    "Enabled"             = ""
    "LastLogin"           = ""
    "LastPasswrdSet"      = ""
    "ItemCount"           = ""  
    "TotalItemSize"       = ""

    "Organizational Unit" = ""
    "ExchangeVersion"     = ""
    "MailboxLocation"     = ""
    "MailboxType"         = ""
    "FullAccess"          = ""
    "SendAs"              = ""
    "SendOnBehalf"        = ""
}

#region Users
Write-Host ""
Write-Host "Getting all mailboxes from Exchange"
$Users = Get-Mailbox -ResultSize Unlimited
Write-Host "- Found $($Users.Count) mailboxes" -ForegroundColor Yellow
Write-Host "- Getting all information from mailboxes"


$Counter =0
ForEach ($User in $users){
        $Counter++
        Write-Progress -Activity 'Processing Users' -PercentComplete (($counter / $Users.count) * 100)
        Start-Sleep -Milliseconds 200

        $Userinformation = Get-Mailbox $($User.Alias) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        $UserStatistics = Get-MailboxStatistics $($User.Alias) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        $Userinformation1 = Get-Aduser $userinformation.SamAccountName -properties * -ErrorAction SilentlyContinue

        $RecordUsers."Name" = $UserInformation.Name
        $RecordUsers."SamAccountName" = $Userinformation.SamAccountName
        $RecordUsers."DisplayName" = $UserInformation.DisplayName
        $RecordUsers."PrimarySMTPAddress" = $Userinformation.PrimarySMTPAddress

        $Aliasses = $Userinformation.Emailaddresses | Where-Object { $_ -match '^smtp:' -and $_ -notcontains "SMTP:$($Userinformation.PrimarySmtpAddress)" }
        If ($Aliasses -eq $null){
            $RecordUsers."Aliasses" = "None"
        }
        else {
            $Addresses = $Aliasses.Split(":") |where-Object {$_ -notlike "smtp"}
            $RecordUsers."Aliasses" = $($Addresses) -join (",")
            }

        $RecordUsers."Enabled" = $UserInformation1.Enabled
        $RecordUsers."UPN" = $UserInformation1.UserPrincipalName
        $RecordUsers."LastPasswrdSet" = $UserInformation1.PasswordLastSet
        $RecordUsers."ItemCount" = $UserStatistics.ItemCount
        $RecordUsers."TotalItemSize" = $UserStatistics.TotalItemSize
        $RecordUsers."Lastlogin" = $UserInformation1.LastLogonDate
        $GettingOU = $UserInformation.DistinguishedName.Split(",") | Where-Object { $_ -notlike "CN=*" }
        $Location = $GettingOU[1..($GettingOU.Length + 1)] -join (",")
        $RecordUsers."Organizational Unit" = $Location
        $RecordUsers."ExchangeVersion" = $UserInformation.ExchangeVersion
        $RecordUsers."MailboxLocation" = $UserInformation.Database
        $RecordUsers."MailboxType" = $UserInformation.RecipientTypeDetails

        $MailboxPermissions = Get-MailboxPermission -Identity $($Userinformation.SamAccountName) -DomainController $DomainControllerToUse -Erroraction SilentlyContinue 
        $MailboxFullPermissions = $MailboxPermissions | Where-Object { $_.AccessRights -like "*full*" -and -not ($_.User -match "NT AUTHORITY") }
        $RecordUsers."FullAccess" = $($MailboxFullPermissions.User) -join (",")
        $MailboxSendAsPermissions = $MailboxPermissions | Where-Object { $_.AccessRights -like "*Send-As*" -and -not ($_.User -match "NT AUTHORITY") } 
        $RecordUsers."SendAs" = $($MailboxSendAsPermissions.User) -join (",")
            
        $permissions = $UserInformation.GrantSendOnBehalfTo
        $RecordUsers."SendOnBehalf" = $($Permissions) -Join (",")

        $objRecordUsers = New-Object PSObject -property $RecordUsers
        $TableUsers += $objRecordUsers

        $TableUsers | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-OnPremExchange-Users-$($Env:USERDOMAIN).csv" -NoTypeInformation

        }

        Write-Progress -Activity "Processing Users" -Status "Ready" -Completed

Write-Host "- Processed $($Users.Count) Mailboxes and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-OnPremExchange-Users-$($Env:USERDOMAIN).csv" -ForegroundColor Green
Write-Host ""
Write-Host "Getting all contacts"
$Contacts = Get-Contact
Write-Host "- Found $($Contacts.Count) Contacts" -ForegroundColor Yellow
Write-Host "- Getting all information from Contacts"
$ContactPerson = @()
ForEach ($Contact in $Contacts){
    $Object = Get-Contact $($Contact.Name) 
    $ContactPerson += $Object
}
$ContactPerson | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-Contacts.csv"

Write-Host "- Processed $($Contacts.Count) Contacts and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-Contacts.csv" -ForegroundColor Green
Write-Host ""
#endregion Users

### GROUPS INVENTORY
Write-Host "Getting all Distribution Groups"
$Groups = Get-DistributionGroup
Write-Host "- Found $($Groups.Count) Distribution Groups"
$GroupInformation= @()
ForEach ($Group in $Groups){
    $object = Get-DistributionGroup $($Group.Name) 
    $GroupInformation += $Object
}
$GroupInformation | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-DistributionGroups.csv"
Write-Host "- Processed $($Groups.Count) Distribution Groups and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-DistributionGroups.csv" -ForegroundColor Green
Write-Host "- Getting all members from Distribution Groups"
$TableGroups = @()
$RecordMembers = [ordered]@{
    "GroupName" = ""
    "DisplayName" = ""
    "PrimarySMTPAddress" = ""
    "Members" = ""
}
ForEach ($Group in $Groups){
    $RecordMembers."GroupName" = $Group.name
    $RecordMembers."DisplayName" = $Group.DisplayName
    $RecordMembers."PrimarySMTPAddress" = $Group.PrimarySMTPAddress

    $Members = Get-DistributionGroupMember $($Group.Name)
    if (!($Members)){
        $RecordMembers."Members" = ""
    }
    Else {
        $RecordMembers."Members" = $($Members.Name) -Join (",")
    }
    $objRecordMembers = New-Object PSObject -property $RecordMembers
    $TableGroups += $objRecordMembers

    $TableGroups | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-DistributionGroupMembers.csv" -NoTypeInformation
}
Write-Host "- Processed $($Groups.Count) distribution groups with there members and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-DistributionGroupMembers.csv" -ForegroundColor Green

#region settings
Write-Host ""
Write-Host "Getting all transport rules"
$Rules = Get-TransportRule
Write-Host "- Found $($Rules.Count) transport rules" -ForegroundColor Yellow
Write-Host "- Exporting all transport rules"
$Transport = @()
ForEach ($Rule in $Rules){
    $Object = Get-TransportRule $($Rule.Name) 
    $Transport += $Object
       
}
$Transport | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-TransportRules.csv"
Write-Host "- Processed $($Rules.Count) transport rules and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-TransportRules.csv" -ForegroundColor Green
Write-Host ""

Write-Host "Getting all Send Connectors"
$sendConnectors = Get-SendConnector
Write-Host "- Found $($sendConnectors.Count) send connectors" -ForegroundColor Yellow
Write-Host "- Exporting all send connectors"
$Send= @()
ForEach ($sendConnector in $sendConnectors){
    $object = Get-SendConnector $($sendConnector.Name) 
    $Send += $Object
}
$Send | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-SendConnectors.csv"
Write-Host "- Processed $($sendConnectors.Count) sent connectors and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-SendConnectors.csv" -ForegroundColor Green
Write-Host ""

Write-Host "Getting all Receive Connectors"
$ReceiveConnectors = Get-ReceiveConnector
Write-Host "- Found $($ReceiveConnectors.Count) receive connectors" -ForegroundColor Yellow
Write-Host "- Exporting all receive connectors"
$Receive = @()
ForEach ($ReceiveConnector in $ReceiveConnectors){
    $Object = Get-ReceiveConnector $($ReceiveConnector.Name) 
    $Receive += $Object
}
$Receive | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-ReceiveConnectors.csv"
Write-Host "- Processed $($ReceiveConnectors.Count) receive connectors and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-ReceiveConnectors.csv" -ForegroundColor Green
Write-Host ""

Write-Host "Getting all accepted domains"
$AcceptedDomains = Get-AcceptedDomain
Write-Host "- Found $($AcceptedDomains.Count) accepted domains" -ForegroundColor Yellow
Write-Host "- Exporting all accepted domains"
$Domain = @()
ForEach ($AcceptedDomain in $AcceptedDomains){
    $Object = Get-AcceptedDomain $($AcceptedDomain.Name) 
    $Domain += $Object
}
$Domain | Export-CSV "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-AcceptedDomains.csv"
Write-Host "- Processed $($AcceptedDomains.Count) accepted domains and exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-AcceptedDomains.csv" -ForegroundColor Green
Write-Host ""
#endregion settings


### VIRTUAL DIRECTORYS
Write-host "Getting information about ActiveSync Virtual Directory"
Get-ActiveSyncVirtualDirectory | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-ActiveSync.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-ActiveSync.txt" -ForegroundColor Green

Write-host "Getting information about Autodiscover Virtual Directory"
Get-AutodiscoverVirtualDirectory | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-AutoDiscover.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-AutoDiscover.txt" -ForegroundColor Green

Write-host "Getting information about OWA Virtual Directory"
Get-OWAVirtualDirectory | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-OWA.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-OWA.txt" -ForegroundColor Green

Write-host "Getting information about ECP Virtual Directory"
Get-ECPVirtualDirectory | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-ECP.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-ECP.txt" -ForegroundColor Green

Write-host "Getting information about WebServices Virtual Directory"
Get-WebServicesVirtualDirectory | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-WebServices.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-WebServices.txt" -ForegroundColor Green

Write-host "Getting information about OfflineAddressBook Virtual Directory"
Get-OABVirtualDirectory  | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-OfflineAddressBook.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-OfflineAddressBook.txt" -ForegroundColor Green

Write-host "Getting information about MAPI Virtual Directory"
Get-MapiVirtualDirectory | fl | out-File "$OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-Mapi.txt"
Write-Host "- Exported to $OUTPUTPATH\$(Get-Date -Format "yyyy-MM-dd")-VD-Mapi.txt" -ForegroundColor Green
Write-Host ""
Remove-PSSession -ComputerName $ExchangeServer
Write-Host "All sessions to Exchange server are closed"

Start-Process explorer.exe $Outputpath

Write-Host "Script end time: $(Get-Date)"
Write-Host "##############################################################################################################" -ForegroundColor Green