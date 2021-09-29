<#
* WARNING * * WARNING * * WARNING * * WARNING * * WARNING * * WARNING * * WARNING * 
This script is designed to delete objects from your Active Directory, Azure Active Directory, Exchange, Intune and SQL. Use at your own risk. 
I provide no warranty or support and accept no responsibility. You should never run random scripts off the internet without understanding their content. 

***

Random reference articles that might help understand components
https://github.com/paulomarquesc/AzureRmStorageTable/issues/61 
https://www.ciraltos.com/write-data-from-powershell-to-azure-table-storage/
https://blog.roostech.se/posts/create-and-use-azure-table-storage-with-powershell/
https://github.com/Microsoft/Intune-PowerShell-SDK

.SYNOPSIS

Audit of AD, AAD user and computer accounts, cleaning of stale objects, data collection for reporting. 

.DESCRIPTION

Script runs several functions that will audit user and computer accounts, upload the data to an Azure Storage table and synchronize to Power BI for reporting. 
The same data is then used to disable and delete user and computer accounts not logged into for 30-60 days. 

.INPUTS

Refer to individual functions. 

.OUTPUTS

Refer to individual functions.

.NOTES
Before this process can work as expected, you need to have created several things in your environment. Apologies this is not a solve-all type of script, but I built on this over time. 
1. I prefer to run this from Azure Automation, so you would need an Automation account setup with a Hybrid Runbook Worker.
2. You need the accounts for your Hybrid Runbook Worker and RunAs account to have the relevant permissions for the actions you want to perform
3. You need an on-prem AD security group called 'SEC - User - No Disable'. Members of this group will not be disabled by this clean up process. Useful for very important accounts that don't get used often. 
4. You need an on-prem AD security group called 'SEC - User - No Delete'. Members of this group will not be deleted by this clean up process. Useful for accounts you need to keep but aren't being used currently, like staff on leave. 
5. An Azure Storage Account that contains the Table names defined in the 'Enter Table Storage location data' area below
6. You'll need to go through the code, especially the Initialisations and Declarations sections, and update variables/commands with your relevant data like server names, Azure resource names, etc. Sorry, I didn't put them all at the top yet like they should be. 

I am not an expert, you will most likely find things I've done badly. Feel free to make improvements. 

#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Import Modules & Snap-ins
$Modules = @('PSWriteHTML'
            'ExchangeOnlineManagement'
            'az'
            'AzureAD'
            'AzTable'
            'Microsoft.Graph.Intune'
            'sqlserver')

# Intune Module: https://github.com/microsoft/Intune-PowerShell-SDK
# EXO Module: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module

foreach ($Module in $Modules) {
    try {
        import-module -name $Module -ErrorAction stop
        write-verbose "Imported $Module"
    }
    catch {
        install-module -name $Module -allowClobber -Force -repository PSGallery
        import-module -name $Module
        write-verbose "Installed and Imported $Module"
    }
}

<# When testing a runbook, verbose messages are not displayed even if the runbook is configured to log verbose records. 
To display verbose messages while testing a runbook, you must set the $VerbosePreference variable to Continue. 
With that variable set, verbose messages will be displayed in the Test Output Pane of the Management Portal.
#>
#$VerbosePreference = 'Continue'    #UnComment this line if you're running the script in the Test Pane

<#region ##### If you run interactively on your PC you need this authentication section only   #####

$cred = get-credential #this should be a local on-prem AD account
$EXLocal = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://yourmailserver/PowerShell/ -Authentication Kerberos
Import-PSSession $EXLocal -AllowClobber

$creds = get-credential #this should be an O365 global admin account or relevant RBAC, may or may not be the same as $cred. 
$EXO = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic –AllowRedirection
Import-PSSession $EXO -AllowClobber

Connect-AzureAD -Credential $creds

# Login for Intune module
Connect-MSGraph -PSCredential $creds

# If you're running interactively you'll need to grab the SAS key for the storage account from Azure portal and use it here. Dont save it in plain text permanently. 
$sasToken = 'Put the SAS key in here'
#endregion #>


#region ##### If you run in Azure Automation you need this authentication section only   #####
write-verbose 'Create Exchange local session'
$EXLocal = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://yourmailserver/PowerShell/ -Authentication Kerberos
Import-Module (Import-PSSession -Session $EXLocal -DisableNameChecking -AllowClobber) -Global

write-verbose 'Creating Exchange Online session'
$tenant = “yourtenant.onmicrosoft.com”
$runAsConnection = Get-AutomationConnection -Name 'AzureRunAsConnection'
Connect-ExchangeOnline -Organization $tenant -appid $runAsConnection.ApplicationId -CertificateThumbprint $runAsConnection.CertificateThumbprint

write-verbose 'Connecting to Azure AD'
Connect-AzureAD -TenantId 'yourtenantGUID' -ApplicationId  $runAsConnection.ApplicationId -CertificateThumbprint $runAsConnection.CertificateThumbprint

# Login for Intune module # https://oofhours.com/2019/11/29/app-based-authentication-with-intune/
# Waiting for certificate authentication to be added to the module: https://github.com/microsoft/Intune-PowerShell-SDK/pull/73
#Note/caution - there is a currently a bug in the module, when you run this from Azure Automation it times out after 60 minutes. Have a support case open with MS. See: https://github.com/microsoft/Intune-PowerShell-SDK/issues/115
write-verbose 'Connecting to MS Graph for Intune'
$MyAppCredential = Get-AutomationPSCredential -Name 'yourautomationcredential1' #Caution, this secret expires. Check the app registration. 
$authority = “https://login.windows.net/$tenant”
$clientId = $runAsConnection.ApplicationId
$clientSecret = $MyAppCredential.GetNetworkCredential().Password
Update-MSGraphEnvironment -AppId $clientId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $ClientSecret -Quiet

# We Need the Storage Account SAS key also. In Az Automation it comes from the Credential Manager.
write-verbose 'Getting SAS token for Storage account'
$MySASCredential = Get-AutomationPSCredential -Name 'yourautomationcredential2'
$sasToken = $MySASCredential.GetNetworkCredential().Password
#endregion
#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Any Global Declarations go here
# 
write-verbose 'Setting all our required variables'
$EXLocal = get-pssession | where {$_.computername -like 'yourmailserver' -and $_.Availability -like 'Available'}
$EXO = get-pssession | where {$_.computername -like 'outlook.office365.com' -and $_.Availability -like 'Available'}
$TodaysDate = get-date -format 'dd-MM-yyyy'
$MOTD = ((iwr http://pastebin.com/raw/dWTSRyr9 -UseBasicParsing).content.split("`n").trim() | ? {!$_.startswith('#')} | random) #just for fun - star wars quote generator. 
[System.Collections.ArrayList]$DisabledUserTable = @()
[System.Collections.ArrayList]$DeletedUserTable = @()
[System.Collections.ArrayList]$DisabledComputerTable = @()
[System.Collections.ArrayList]$DeletedComputerTable = @()
[System.Collections.ArrayList]$DisabledAADComputerTable = @()
[System.Collections.ArrayList]$DeletedAADComputerTable = @()

# Enter Table Storage location data 
$storageAccountName = 'yourstorageaccount'
$RG = 'yourstorageaccountresourcegroup'
$tableName = 'ADUsers1'
$UserMetricsTableName = 'ADUserMetrics1'
$ComputerTableName = 'ADComputers1'
$ComputerMetricsTableName = 'ADComputerMetrics1'
$AADComputersTableName = 'AADComputers1'
$AADComputerMetricsTableName = 'AADComputerMetrics1'
$partitionKey = 'PartitionKey1'

# Connect to Azure Table Storage
$storageCtx = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $sasToken
$table = (Get-AzStorageTable -Name $tableName -Context $storageCtx).CloudTable
$UserMetricsTable = (Get-AzStorageTable -Name $UserMetricsTableName -Context $storageCtx).CloudTable
$ComputersTable = (Get-AzStorageTable -Name $ComputerTableName -Context $storageCtx).CloudTable
$ComputerMetricsTable = (Get-AzStorageTable -Name $ComputerMetricsTableName -Context $storageCtx).CloudTable
$AADComputersTable = (Get-AzStorageTable -Name $AADComputersTableName -Context $storageCtx).CloudTable
$AADComputerMetricsTable = (Get-AzStorageTable -Name $AADComputerMetricsTableName -Context $storageCtx).CloudTable

# Which SQL servers are we going to clean users out of? 
$SQLServers = @('hostname1\instance1'
                'hostname2'
                )

#-----------------------------------------------------------[Functions]------------------------------------------------------------
#region Functions are in here
function Get-MailboxLastLogon {
    <#
    .SYNOPSIS
    Retrieves the Exchange mailboxes last logon value for the specified user.  
    .DESCRIPTION
    The mailbox last logon value is important for shared mailboxes where the user account shows no login for a long period of time, but the
    mailbox has active daily logins from other user accounts. 
    .EXAMPLE
    Get-MailboxLastLogon -useremail 'han.solo@yourdomain.com'      #Gets the last login date/time value for Han Solo's mailbox in Exchange. 
    .PARAMETER
    -UserEmail       Mandatory
    .INPUTS
    Only input accepted is UserEmail - which mailbox do you want to retrieve the attribute from.
    .OUTPUTS
    Date/Time value will be returened. Best captured as a variable and piped to other commands. 
    .NOTES

#>
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$UserEmail
    )
    try {
        $MXLastLogoncheck = invoke-command -session $EXLocal -scriptblock { Get-MailboxStatistics $Using:UserEmail } -ErrorAction stop
        $MXLastLogon =  $MXLastLogoncheck.lastlogontime 
        $MXLastLogon
    }
    catch {
        $MXLastLogoncheck = invoke-command -session $EXO -scriptblock { Get-MailboxStatistics $Using:UserEmail } -erroraction silentlycontinue
        if (-not($MXLastLogoncheck)) {  $MXLastLogon = ''   } else { $MXLastLogon = $MXLastLogoncheck.lastlogontime    }
        $MXLastLogon
    }      
} 

#

function Delete-SQLUser {
    #This is only useful if you have added AD user accounts directly to SQL. The preferred method of assigning SQL permission should be via AD group membership. 
    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$UserSAM
    )
    foreach ($SQLServer in $SQLServers) {
        try {
            get-sqllogin -ServerInstance $SQLServer -LoginName "YourDomain\$UserSAM"  -ErrorAction stop | remove-sqllogin -RemoveAssociatedUsers  -Force
            write-verbose "Deleted $UserSAM from $SQLServer"
        }
        catch {
            write-verbose "$UserSAM not in $SQLServer"
        }
    }
}

Function Upload-ADUserInfo {

    <#
        .SYNOPSIS
        Uploads data for recently modified user accounts to Azure Table Storage. 
        .DESCRIPTION
        This function uses the 'all users' variable, filters to the user accounts modified in the last x days and uploads the data for those accounts to the Azure Table. 
        New accounts will have a record created, accounts that already exist in the Table will be updated. 
        .EXAMPLE
        Upload-ADUserInfo -previousdays 30      #Uploads data for all user accounts modified in the last 30 days
        .PARAMETER
        -PreviousDays       Not mandatory, defaults to 1
        .INPUTS
        Only input accepted is PreviousDays - how many days previous do you want to check the WhenModified property for accounts to target with this function. Default is '1'. 
        I'm not convinced this is a reliable way to filter out 'unchanged' user objects to speed up your run, needs more testing. I currently don't use this filter/switch. 
        .OUTPUTS
        All data is uploaded directly to the Azure Storage Table defined in the Global Declarations of this script.
        .NOTES
        Really should rebuild this function and remove its reliance on $AllUsers - it should be a single-user function that gets piped through a foreach at run time. To-do. 
    #>
    
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
            [string]$PreviousDays = '1'
        )
    
    Begin {
    
        $time = (Get-Date).Adddays(-($PreviousDays)) 
    
        write-verbose "Uploading details of AD user accounts modified since $time to Azure Table Storage..."
    
    }
    
    Process {
    
        Try {
            foreach ($User in $users) {
            
            if ($User.whenChanged -gt $Time ) { 
            try { 
    
                            # Some accounts dont have values for certain properties, they will throw errors. Lets fix that with a bogus value.
                            foreach ($p in $User.PSObject.Properties) { if ($Null -eq $p.Value -and $p.TypeNameOfValue -eq 'System.String' -and   $p.IsSettable -eq $true) { $p.Value = ''}} Write-Output $_
    
                            #These next 2 lines in theory could be consolidated into the PropertyArray - to do
                            if ($User.memberof -like "CN=SEC - User - No Disable*") {$User | add-member 'NoDisable' 'True' -force} else {$User | add-member 'NoDisable' 'False' -force} 
                            if ($User.memberof -like "CN=SEC - User - No Delete*") {$User | add-member 'NoDelete' 'True' -force} else {$User | add-member 'NoDelete' 'False' -force} 
    
                    $UserGUID = $user.ObjectGUID 
                    $CurrentRecord = get-azTableRow -table $table -columnName "ObjectGUID" -guidValue ([guid]"$UserGUID")  -operator Equal
    
    
                    $PropertyArray = @{
                        
                        'Name' = $User.Name
                        'GivenName' = if ($user.GivenName) {$user.GivenName} else {''}
                        'Surname' = if ($user.Surname) {$User.Surname} else {''}
                        'DistinguishedName' = $User.DistinguishedName
                        'UserPrincipalName' = $User.UserPrincipalName
                        'SamAccountName' = $User.SamAccountName
                        'Enabled' = $User.enabled
                        'City' = $User.city
                        'State' = $User.state
                        'Company' = $user.Company
                        'Department' = $user.Department
                        'Office' = $User.office
                        'Description' = $User.Description
                        'Title' = $user.Title
                        'EmailAddress' = $user.EmailAddress
                        'LastLogonDate' = if ($Null -ne $User.LastLogonDate) {$User.LastLogonDate } else {''} # https://social.technet.microsoft.com/wiki/contents/articles/22461.understanding-the-ad-account-attributes-lastlogon-lastlogontimestamp-and-lastlogondate.aspx
                        'LockedOut' = $user.LockedOut
                        #'MemberOf' = $user.MemberOf        #Currently doesn't work, need to turn it into CSV instead of array
                        'OfficePhone' = $User.OfficePhone
                        'MobilePhone' = $User.MobilePhone
                        'PasswordExpired' = $User.PasswordExpired
                        'PasswordLastSet' = if ($Null -ne $User.PasswordLastSet) {$User.PasswordLastSet } else {''}
                        'PasswordNeverExpires' = $User.PasswordNeverExpires    
                        'whenChanged' = if ($Null -ne $User.whenChanged) {$User.whenChanged } else {''} #Probably irrelevant, field should always be populated
                        'whenCreated' = if ($Null -ne $User.whenCreated) {$User.whenCreated } else {''} #Probably irrelevant, field should always be populated
                        'ObjectGUID' = $User.ObjectGUID
                        'MailboxLastLogon' = $user.MailboxLastLogon
                        'SkypeEnabled' = if ($User.'msRTCSIP-UserEnabled' -eq $True ) { 'True'  } else { 'False' }
                        'NoDisable' = $User.NoDisable 
                        'NoDelete' = $User.NoDelete 
                        'DateDeleted' = ''
                        'DateDisabled' = if (($User.Enabled -eq $False ) -and ($CurrentRecord.Enabled -eq $True <#Suspect redundant#>)) { get-date } elseif ($User.Enabled -eq $True) { '' } elseif ($User.Enabled -eq $False) { get-date }
                    }
    
            
                    $Username = $user.name 
                    #write-verbose "Uploading data for $Username" #Comment out if not required, fills up logs when you have a high user count
    
                    if ($Null -eq $CurrentRecord) {Add-azTableRow -table $table -partitionKey $partitionKey -rowKey ((New-Guid).Guid) -property $PropertyArray | out-null  }
                    else { 
                            # this could be cleaned up
                            $CurrentRecord.Name = $User.Name
                            $CurrentRecord.GivenName = $User.GivenName
                            $CurrentRecord.Surname = $User.Surname
                            $CurrentRecord.DistinguishedName = $User.DistinguishedName
                            $CurrentRecord.UserPrincipalName = $User.UserPrincipalName
                            $CurrentRecord.SamAccountName = $User.SamAccountName
                            $CurrentRecord.Enabled = $User.Enabled
                            $CurrentRecord.City = $User.City
                            $CurrentRecord.State = $User.State
                            $CurrentRecord.Company = $User.Company
                            $CurrentRecord.Department = $User.Department
                            $CurrentRecord.office = $User.office
                            $CurrentRecord.description = $User.description
                            $CurrentRecord.Title = $User.Title
                            $CurrentRecord.EmailAddress = $User.EmailAddress
                            $CurrentRecord.LastLogonDate = $PropertyArray.lastlogondate
                            $CurrentRecord.LockedOut = $User.LockedOut
                            $CurrentRecord.OfficePhone = $User.OfficePhone
                            $CurrentRecord.MobilePhone = $User.MobilePhone
                            $CurrentRecord.PasswordExpired = $User.PasswordExpired
                            $CurrentRecord.PasswordLastSet = $PropertyArray.passwordlastset
                            $CurrentRecord.PasswordNeverExpires = $User.PasswordNeverExpires
                            $CurrentRecord.whenChanged = $User.whenChanged
                            $CurrentRecord.MailboxLastLogon = $User.MailboxLastLogon
                            $CurrentRecord.SkypeEnabled = $PropertyArray.SkypeEnabled
                            $CurrentRecord.NoDisable = $PropertyArray.NoDisable
                            $CurrentRecord.NoDelete = $PropertyArray.NoDelete
                            if ($PropertyArray.DateDisabled) {$CurrentRecord.DateDisabled = $PropertyArray.DateDisabled}
                            $CurrentRecord | update-aztablerow -table $table  |  out-null    }
                                    
                }
                catch {
                    write-warning "Error updating user $Username"
                    Write-warning "Error: $($_.Exception.Message)"
                }
                }
            } 
       }
    
        Catch {
        Write-warning "Error: $($_.Exception.Message)"
        Break
        }
    }
    
    End {
    
    If ($?) {
    
    write-verbose 'Completed updating AD user info to Azure Table successfully.'
    }
    }
    }


Function Disable-StaleADUser {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DisableDays = '30'
    )

    Begin {

        $DisableTime = (Get-Date).Adddays(-($DisableDays))  
        $NewAccount = (Get-Date).Adddays(-(14)) #used to filter out new accounts just created but not logged into yet
        $LyncOU = "*OU=Your,OU=OU,OU=Path,DC=Your,DC=Domain" #Used to filter a specific OU and apply different process to it
        $EXHealthMailbox = "*CN=Monitoring Mailboxes,CN=Microsoft Exchange System Objects,DC=Your,DC=Domain" #Filter out these mailboxes, you should leave them alone.

        write-verbose "User accounts not logged into since $DisableTime will be disabled."

    }

    Process {
    Try {
        foreach ($User in $users) {

            $Name = $user.displayname
            $JobTitle = $user.Title
            $Dept = $user.department
            $Company = $user.company
            $Description = $User.Description
            $UserSAM = $User.samaccountname
            $LastLogon = $User.lastlogondate
            $LastMBLogon = $User.MailboxLastLogon
            $NoDisable = $User.NoDisable
            
            
            if ($User.LastLogondate -lt $DisableTime -and $User.MailboxLastLogon -lt $DisableTime -and $User.Enabled -eq $True -and $User.DistinguishedName -notlike $LyncOU -and $User.DistinguishedName -notlike $EXHealthMailbox -and $user.whenCreated -lt $NewAccount -and $NoDisable -eq 'False' ) { 
                
                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'SAM Account Name' = $UserSAM
                    'Job Title'  = $JobTitle
                    'Department' = $Dept
                    'Company' = $Company
                    'Description' = $Description
                    'Last Logon' = $LastLogon
                    'Mailbox Last Logon' = $LastMBLogon
                    }
                $DisabledUserTable.add($Obj)       
                
                $UserGUID = $user.ObjectGUID 
                $CurrentRecord = get-azTableRow -table $table -columnName "ObjectGUID" -guidValue ([guid]"$UserGUID")  -operator Equal

                write-verbose "Disabling user account for $UserSAM"
                Set-aduser $user.ObjectGUID -enabled $False -whatif #turn whatif on/off for dry/live runs
                
                $CurrentRecord | add-member 'DateDisabled' (get-date) -force 
                $CurrentRecord | update-aztablerow -table $table  |  out-null 
                }
            elseif ($User.LastLogondate -lt $DisableTime -and $User.MailboxLastLogon -lt $DisableTime -and $User.Enabled -eq $True -and $User.DistinguishedName -like $LyncOU -and $user.whenCreated -lt $NewAccount -and $User.SkypeEnabled -eq $False) {
                #Applies different behaviour to accounts in a specific OU. Can remove this if not needed.      
                $Name = $user.displayname
                $JobTitle = $user.Title
                $Dept = $user.department
                $Company = $user.company
                $Description = $User.Description
                $UserSAM = $User.samaccountname
                $LastLogon = $User.lastlogondate
                $LastMBLogon = $User.MailboxLastLogon               
    
                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'SAM Account Name' = $UserSAM
                    'Job Title'  = $JobTitle
                    'Department' = $Dept
                    'Company' = $Company
                    'Description' = $Description
                    'Last Logon' = $LastLogon
                    'Mailbox Last Logon' = $LastMBLogon
                    }
                $DisabledUserTable.add($Obj)
                
                $UserSAM = $user.SamAccountName
                $UserGUID = $user.ObjectGUID 
                $CurrentRecord = get-azTableRow -table $table -columnName "ObjectGUID" -guidValue ([guid]"$UserGUID")  -operator Equal

                write-verbose "Disabling user account for $UserSAM"
                Set-aduser $user.ObjectGUID -enabled $False -whatif #turn whatif on/off for dry/live runs
                
                $CurrentRecord | add-member 'DateDisabled' (get-date) -force 
                $CurrentRecord | update-aztablerow -table $table  |  out-null 
            }
            }
            write-verbose 'Successfully completed disabling users.'
    } 
    Catch {
        Write-warning -BackgroundColor Red "Error: $($_.Exception.Message)"
        Break
            }
   }
}

###
Function Delete-StaleADUser {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DeleteDays = '60'
    )

Begin {

    $DeleteTime = (Get-Date).Adddays(-($DeleteDays)) 

    write-verbose "User accounts not logged into since $DeleteTime will be deleted."

}

Process {


    Try {
        foreach ($User in $users) {
    
            if ($User.LastLogondate -lt $DeleteTime -and $User.MailboxLastLogon -lt $DeleteTime -and $User.Enabled -eq $False -and $User.NoDelete -eq $False ) {

                $Name = $user.displayname
                $JobTitle = $user.Title
                $Dept = $user.department
                $Company = $user.company
                $Description = $User.Description
                $UserSAM = $User.samaccountname
                $LastLogon = $User.lastlogondate
                $LastMBLogon = $User.MailboxLastLogon               
    
                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'SAM Account Name' = $UserSAM
                    'Job Title'  = $JobTitle
                    'Department' = $Dept
                    'Company' = $Company
                    'Description' = $Description
                    'Last Logon' = $LastLogon
                    'Mailbox Last Logon' = $LastMBLogon
                    }
                $DeletedUserTable.add($Obj) 

                $UserGUID = $user.ObjectGUID 
                $CurrentRecord = get-azTableRow -table $table -columnName "ObjectGUID" -guidValue ([guid]"$UserGUID")  -operator Equal

                write-verbose "Deleting user account for $UserSAM"
                # Delete the user here
                try {

                    Remove-ADUser -Identity $UserGUID -confirm:$false -ErrorAction Stop -WhatIf #turn whatif on/off for dry/live runs
                
                } 
                catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {write-output "We already deleted $UserSAM " }
                catch {
                
                    Remove-ADObject -Identity $UserGUID -Recursive -confirm:$false -WhatIf #turn whatif on/off for dry/live runs
                
                }

                #Need to delete the user from SQL also, if they exist there
                Delete-SQLUser -UserSAM $UserSAM #enable/disable this for dry/live runs
                
                $CurrentRecord | add-member 'DateDeleted' (get-date) -force
                $CurrentRecord | update-aztablerow -table $table  |  out-null 
            }
        }
        write-verbose 'Successfully completed deleting users.'
} 
Catch {

    Write-warning "Error: $($_.Exception.Message) for $UserSAM"

        }
}
}


###
function New-TableColumn {
    <#
    .SYNOPSIS
    Uploads a new table column to the users or computers storage tables
    
    .DESCRIPTION
    If you want to add a new column to the table, you need to add it as a record (key/value pair) against every row in the table. 
    
    .PARAMETER NewColumnName
    What is the name of the new column you want to create? No spaces allowed. 
    
    .PARAMETER UsersOrComputers
    Do you want to create the column in the users table (ADUsers1) or the computers table (ADComputers1). Needs to be expanded for other tables. 
    
    .EXAMPLE
    new-tablecolumn -NewColumnName 'DeathStar' -UsersOrComputers 'Computers'        #Creates the DeathStar column against every record in the ADComputers1 table. 
    
    .NOTES
    Function not currently setup to add any columns to the Metrics tables. Also remember if you add a new column to existing records, adjust your function so new records are also created with the new columns. 
    #>
    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$NewColumnName,
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("ADUsers1","ADComputers1","AADComputers1")]
        [string]$TargetTableName
    )
    if ($TargetTableName -eq 'ADUsers1') {
        foreach ($User in $Users) {
            $UserGUID = $user.ObjectGUID 
            $CurrentRecord = get-azTableRow -table $table -columnName "ObjectGUID" -guidValue ([guid]"$UserGUID")  -operator Equal
        
            $Username = $user.name 
            write-verbose "Creating new column for $Username"
    
            $CurrentRecord | add-member "$NewColumnName" '' -force
            $CurrentRecord | update-aztablerow -table $table  |  out-null
        }   
    }
    elseif ($TargetTableName -eq 'ADComputers1') {
        foreach ($Computer in $Computers) {
            $ComputerGUID = $Computer.ObjectGUID 
            $CurrentRecord = get-azTableRow -table $ComputersTable -columnName "ObjectGUID" -guidValue ([guid]"$ComputerGUID")  -operator Equal
        
            $Computername = $Computer.name 
            write-verbose "Creating new column for $Computername"
    
            $CurrentRecord | add-member "$NewColumnName" '' -force
            $CurrentRecord | update-aztablerow -table $ComputersTable  |  out-null
        }
    }
    elseif ($TargetTableName -eq 'AADComputers1') {
        foreach ($AADComputer in $AADComputers) {
            $ComputerID = $AADComputer.ObjectID 
            $CurrentRecord = get-azTableRow -table $AADComputersTable -columnName "ObjectID" -Value "$ComputerID"  -operator Equal
        
            $AADComputername = $AADComputer.displayname 
            write-verbose "Creating new column for $AADComputername"
    
            $CurrentRecord | add-member "$NewColumnName" '' -force
            $CurrentRecord | update-aztablerow -table $AADComputersTable  |  out-null
    }

}
}




function Upload-UserMetrics {

    $PropertyArray = @{
                
        'TotalUsers' = ($users).count
        'EnabledTrue' = ($Users | where {$_.Enabled -eq $True}).count
        'EnabledFalse' = ($Users | where {$_.Enabled -eq $False}).count
        'LockedOutTrue' = ($Users | where {$_.LockedOut -eq $True}).count
        'LockedOutFalse' = ($Users | where {$_.LockedOut -eq $False}).count
        'NoDisableTrue' = ($Users | where {$_.NoDisable -eq $True} | measure-object).count
        'NoDisableFalse' = ($Users | where {$_.NoDisable -eq $False}).count
        'NoDeleteTrue' = ($Users | where {$_.NoDelete -eq $True}).count
        'NoDeleteFalse' = ($Users | where {$_.NoDelete -eq $False}).count
        'PasswordExpiredTrue' = ($Users | where {$_.PasswordExpired -eq $True}).count
        'PasswordExpiredFalse' = ($Users | where {$_.PasswordExpired -eq $False}).count
        'PasswordNeverExpiresTrue' = ($Users | where {$_.PasswordNeverExpires -eq $True}).count
        'PasswordNeverExpiresFalse' = ($Users | where {$_.PasswordNeverExpires -eq $False}).count
    }
    
    Add-azTableRow -table $UserMetricsTable -partitionKey $partitionKey -rowKey ((New-Guid).Guid) -property $PropertyArray | out-null
}

Function Upload-ADComputerInfo {

    <#
        .SYNOPSIS
        Uploads data for computer accounts to Azure Table Storage. 
        .DESCRIPTION
        This function uses the 'all computers' variable, filters to the computer accounts modified in the last x days and uploads the data for those accounts to the Azure Table. 
        New accounts will have a record created, accounts that already exist in the Table will be updated. 
        .EXAMPLE
        Upload-ADComputerInfo -previousdays 30      #Uploads data for all computer accounts modified in the last 30 days
        .PARAMETER
        -PreviousDays       Not mandatory, defaults to 1
        .INPUTS
        Only input accepted is PreviousDays - how many days previous do you want to check the WhenModified property for accounts to target with this function. Default is '1'.
        .OUTPUTS
        All data is uploaded directly to the Azure Storage Table defined in the Global Declarations of this script.
        .NOTES
        Really should rebuild this function and remove its reliance on $AllADComputers - it should be a single-user function that gets piped through a foreach at run time. To-do. 
    #>
    
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
            [string]$PreviousDays = '1'
        )
    
    Begin {
    
        $time = (Get-Date).Adddays(-($PreviousDays)) 
    
        write-verbose "Uploading details of AD computer accounts modified since $time to Azure Table Storage..."
    
    }
    
    Process {
    Try {
        foreach ($Computer in $Computers) {
        
         if ($Computer.whenChanged -gt $Time ) { 
           try { 
    
                        # Some accounts dont have values for certain properties, they will throw errors. Lets fix that with a bogus value.
                        foreach ($p in $Computer.PSObject.Properties) { if ($Null -eq $p.Value -and $p.TypeNameOfValue -eq 'System.String' -and   $p.IsSettable -eq $true) { $p.Value = ''}} Write-Output $_
                        
                        $ComputerGUID = $Computer.ObjectGUID 
                        $CurrentRecord = get-azTableRow -table $ComputersTable -columnName "ObjectGUID" -guidValue ([guid]"$ComputerGUID")  -operator Equal

                $PropertyArray = @{
                    
                    'Name' = $Computer.Name
                    'DNSHostName' = if ($Null -ne $Computer.DNSHostName) {$Computer.DNSHostName } else {''}
                    'IPv4Address' = if ($Null -ne $Computer.IPv4Address) {$Computer.IPv4Address } else {''}
                    'DistinguishedName' = $Computer.DistinguishedName
                    'Enabled' = $Computer.enabled
                    'Description' = $Computer.Description
                    'LastLogonDate' = if ($Null -ne $Computer.LastLogonDate) {$Computer.LastLogonDate } else {''} # https://social.technet.microsoft.com/wiki/contents/articles/22461.understanding-the-ad-account-attributes-lastlogon-lastlogontimestamp-and-lastlogondate.aspx
                    'LockedOut' = $Computer.LockedOut
                    'PasswordExpired' = $Computer.PasswordExpired
                    'PasswordLastSet' = if ($Null -ne $Computer.PasswordLastSet) {$Computer.PasswordLastSet } else {''}
                    'whenChanged' = if ($Null -ne $User.whenChanged) {$Computer.whenChanged } else {''} #Probably irrelevant, field should always be populated
                    'whenCreated' = if ($Null -ne $Computer.whenCreated) {$Computer.whenCreated } else {''} #Probably irrelevant, field should always be populated
                    'ObjectGUID' = $Computer.ObjectGUID
                    'OperatingSystem' = $Computer.OperatingSystem
                    'OperatingSystemVersion' = $Computer.OperatingSystemVersion
                    'DateDeleted' = ''
                    'DateDisabled' = if (($Computer.Enabled -eq $False ) -and ($CurrentRecord.Enabled -eq $True)) { get-date } elseif ($Computer.Enabled -eq $True) { '' }
                }
               
                $Hostname = $Computer.name 
                #write-verbose "Uploading data for $Hostname" #Comment out if not required, fills up logs
    
                if ($Null -eq $CurrentRecord) {Add-azTableRow -table $ComputersTable -partitionKey $partitionKey -rowKey ((New-Guid).Guid) -property $PropertyArray | out-null  }
                else { 
                    $CurrentRecord.Name = $Computer.Name
                    $CurrentRecord.DNSHostName = $PropertyArray.DNSHostName
                    $CurrentRecord.IPv4Address = $PropertyArray.IPv4Address
                    $CurrentRecord.DistinguishedName = $Computer.DistinguishedName
                    $CurrentRecord.Enabled = $Computer.Enabled
                    $CurrentRecord.description = $Computer.description
                    $CurrentRecord.LastLogonDate = $PropertyArray.lastlogondate 
                    $CurrentRecord.LockedOut = $Computer.LockedOut
                    $CurrentRecord.PasswordExpired = $Computer.PasswordExpired
                    $CurrentRecord.PasswordLastSet = $PropertyArray.passwordlastset 
                    $CurrentRecord.whenChanged = $Computer.whenChanged
                    $CurrentRecord.OperatingSystem = $PropertyArray.OperatingSystem
                    $CurrentRecord.OperatingSystemVersion = $PropertyArray.OperatingSystemVersion
                    if ($PropertyArray.DateDisabled) {$CurrentRecord.DateDisabled = $PropertyArray.DateDisabled}
                    $CurrentRecord | update-aztablerow -table $ComputersTable  |  out-null    }
                
            }
            catch {
                write-warning "Error updating computer $Hostname"
                Write-warning "Error: $($_.Exception.Message)"
            }
            }
           } 
        
    
    }
    
    Catch {
    
    Write-warning "Error: $($_.Exception.Message)"
    
    Break
    
    }
    
    }
    
    End {
    
    If ($?) {
    
        write-verbose 'Completed updating AD computer info to Azure Table successfully.'
    
    
    }
    
    }
    
    }
##
function Upload-ComputerMetrics {

    $PropertyArray = @{
                
        'TotalComputers' = ($Computers).count
        'EnabledTrue' = ($Computers | where {$_.Enabled -eq $True}).count
        'EnabledFalse' = ($Computers | where {$_.Enabled -eq $False}).count
        'WindowsServer2000' = ($Computers | where {$_.OperatingSystem -like 'Windows 2000 Server'}).count
        'WindowsServer2003' = ($Computers | where {$_.OperatingSystem -like 'Windows Server 2003'}).count
        'WindowsServer2008' = ($Computers | where {$_.OperatingSystem -like 'Windows* Server* 2008*'}).count
        'WindowsServer2012' = ($Computers | where {$_.OperatingSystem -like 'Windows Server 2012*'}).count
        'WindowsServer2016' = ($Computers | where {$_.OperatingSystem -like 'Windows Server 2016*'}).count
        'WindowsServer2019' = ($Computers | where {$_.OperatingSystem -like 'Windows Server 2019*'}).count
        'Windows7' = ($Computers | where {$_.OperatingSystem -like 'Windows 7*'}).count
        'Windows10' = ($Computers | where {$_.OperatingSystem -like 'Windows 10*'}).count
    }
    
    Add-azTableRow -table $ComputerMetricsTable -partitionKey $partitionKey -rowKey ((New-Guid).Guid) -property $PropertyArray | out-null
}

##

Function Disable-StaleADComputer {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DisableDays = '30'

    )

    Begin {

        $DisableTime = (Get-Date).Adddays(-($DisableDays))  

        write-verbose "Computer accounts not logged into since $DisableTime will be disabled."

    }

    Process {
    Try {
        foreach ($Computer in $Computers) {

            $Name = $Computer.name
            $ComputerGUID = $Computer.ObjectGUID 
            $Description = $Computer.Description
            $LastLogon = $Computer.lastlogondate
            $ComputerinAAD = $AADComputers | ? {$_.Deviceid -like "$ComputerGUID"}
            $AADLastLogon = $ComputerinAAD.ApproximateLastLogonTimeStamp            
            
            if ( ($Computer.LastLogondate -lt $DisableTime -and $Computer.Enabled -eq $True -and $Null -eq $ComputerinAAD) -or ($Computer.LastLogondate -lt $DisableTime -and $Computer.Enabled -eq $True -and $AADLastLogon -lt $DisableTime) ) { 
                
                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'Description' = $Description
                    'Last Logon' = $LastLogon
                    'AAD Last Logon' = $AADLastLogon
                    }
                $DisabledComputerTable.add($Obj)       
                
                $CurrentRecord = get-azTableRow -table $ComputersTable -columnName "ObjectGUID" -guidValue ([guid]"$ComputerGUID")  -operator Equal

                write-verbose "Disabling computer account for $Name"
                Set-adcomputer $Computer.ObjectGUID -enabled $False -whatif
                
                $CurrentRecord | add-member 'DateDisabled' (get-date) -force 
                $CurrentRecord | update-aztablerow -table $ComputersTable  |  out-null 
                }
            }
            write-verbose 'Successfully completed disabling computers.'
    } 
    Catch {

        Write-wa  -BackgroundColor Red "Error: $($_.Exception.Message)"

        Break
            }
    }
}
##

Function Delete-StaleADComputer {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DeleteDays = '60'
    )

Begin {

    $DeleteTime = (Get-Date).Adddays(-($DeleteDays)) 

    write-verbose "Computer accounts not logged into since $DeleteTime will be deleted."

}

Process {
    Try {
        foreach ($Computer in $Computers) {

            $Name = $Computer.name
            $ComputerGUID = $Computer.ObjectGUID 
            $Description = $Computer.Description
            $LastLogon = $Computer.lastlogondate
            $ComputerinAAD = $AADComputers | ? {$_.Deviceid -like "$ComputerGUID"}
            $AADLastLogon = $ComputerinAAD.ApproximateLastLogonTimeStamp
    
            if ( ($Computer.LastLogondate -lt $DeleteTime -and $Computer.Enabled -eq $False -and $Null -eq $ComputerinAAD) -or ($Computer.LastLogondate -lt $DeleteTime -and $Computer.Enabled -eq $False -and $AADLastLogon -lt $DeleteTime) ){

                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'Description' = $Description
                    'Last Logon' = $LastLogon
                    'AAD Last Logon' = $AADLastLogon
                    }
                $DeletedComputerTable.add($Obj) 

                $ComputerGUID = $Computer.ObjectGUID 
                $CurrentRecord = get-azTableRow -table $ComputersTable -columnName "ObjectGUID" -guidValue ([guid]"$ComputerGUID")  -operator Equal

                write-verbose "Deleting computer account for $Name"
                # Delete the computer here
                try {

                    Remove-ADComputer -Identity $ComputerGUID -confirm:$false -ErrorAction Stop -whatif
                
                } catch {
                
                    Remove-ADObject -Identity $ComputerGUID -Recursive -confirm:$false -whatif
                
                }
                
                $CurrentRecord | add-member 'DateDeleted' (get-date) -force
                $CurrentRecord | update-aztablerow -table $ComputersTable  |  out-null 
            }
        }
        write-verbose 'Successfully completed deleting computers.'
} 
Catch {

    Write-warning -BackgroundColor Red "Error: $($_.Exception.Message)"

    Break

        }
}
}

##

Function Upload-AADComputerInfo {

    <#
        .SYNOPSIS
        Uploads data for Azure AD computer accounts to Azure Table Storage. 
        .DESCRIPTION
        This function uses the 'all AAD computers' variable to get data from all AAD device objects and uploads the data for those accounts to the Azure Table. 
        New accounts will have a record created, accounts that already exist in the Table will be updated. 
        The function uses three seperate powershell cmdlets to gather information and compiles them all back into the $AADComputer object to be used in later commands. 
        .EXAMPLE
        Upload-AADComputerInfo 
        .PARAMETER
        No parameters for this function
        .INPUTS
        Null
        .OUTPUTS
        All data is uploaded directly to the Azure Storage Table defined in the Global Declarations of this script.
        .NOTES
        Really should rebuild this function and remove its reliance on $AllADComputers - it should be a single-user function that gets piped through a foreach at run time. To-do. 
    #>
    
        [CmdletBinding()]
        Param (
        )
    
    Begin {
    
        write-verbose "Uploading details of Azure AD computer accounts to Azure Table Storage..."
    
    }
    
    Process {
    Try {
        foreach ($AADComputer in $AADComputers) {
        
           try { 
            
            $DeviceID = $AADComputer.DeviceID
            $Objectid = $AADComputer.ObjectId

            $AADRegisteredOwner = Get-AzureADDeviceRegisteredOwner -OBJECTID $Objectid
            $AADComputer | add-member 'AADRegisteredOwnerUPN' $AADRegisteredOwner.UserPrincipalName -force
            $AADComputer | add-member 'AADRegisteredOwnerEnabled' $AADRegisteredOwner.AccountEnabled -force

            $IntuneDevice = Get-IntuneManagedDevice | ? {$_.azureADDeviceId -like $DeviceID}
            $IntuneManaged = if ($IntuneDevice ) {'True' } else {'False'}
            $AADComputer | add-member 'managedDeviceOwnerType' $IntuneDevice.managedDeviceOwnerType -force
            $AADComputer | add-member 'enrolledDateTime' $IntuneDevice.enrolledDateTime -force
            $AADComputer | add-member 'lastSyncDateTime' $IntuneDevice.lastSyncDateTime -force
            $AADComputer | add-member 'deviceEnrollmentType' $IntuneDevice.deviceEnrollmentType -force
            $AADComputer | add-member 'isEncrypted' $IntuneDevice.isEncrypted -force
            $AADComputer | add-member 'model' $IntuneDevice.model -force
            $AADComputer | add-member 'manufacturer' $IntuneDevice.manufacturer -force
            $AADComputer | add-member 'serialNumber' $IntuneDevice.serialNumber -force
            $AADComputer | add-member 'IntuneManaged' $IntuneManaged -force

            # Some accounts dont have values for certain properties, they will throw errors. Lets fix that with a bogus value.
            foreach ($p in $AADComputer.PSObject.Properties) { if ($Null -eq $p.Value -and $p.TypeNameOfValue -eq 'System.String' -and   $p.IsSettable -eq $true) { $p.Value = ''}} Write-Output $_
                remove-variable PropertyArray -erroraction silentlycontinue
                
                $CurrentRecord = get-azTableRow -table $AADComputersTable -columnName "ObjectID" -Value "$Objectid"  -operator Equal
                
                $PropertyArray = @{
                    
                    'Name' = if ($Null -ne $AADComputer.DisplayName) {$AADComputer.DisplayName } else {''}
                    'ObjectID' = $AADComputer.ObjectID
                    'DeviceID' = $AADComputer.DeviceID
                    'AccountEnabled' =  if ($Null -ne $AADComputer.AccountEnabled) {$AADComputer.AccountEnabled } else {''}
                    'ApproximateLastLogonTimeStamp' =  if ($Null -ne $AADComputer.ApproximateLastLogonTimeStamp) {$AADComputer.ApproximateLastLogonTimeStamp } else {''}
                    'DeviceOSType' =   if ($Null -ne $AADComputer.DeviceOSType) {$AADComputer.DeviceOSType } else {''}
                    'DeviceOSVersion' =  if ($Null -ne $AADComputer.DeviceOSVersion) {$AADComputer.DeviceOSVersion } else {''}
                    'DeviceTrustType' =  if ( $AADComputer.DeviceTrustType) {$AADComputer.DeviceTrustType } else {''}
                    'DirSyncEnabled' = if ($Null -ne $AADComputer.DirSyncEnabled) {$AADComputer.DirSyncEnabled } else {''}
                    'IsCompliant' =  if ($AADComputer.IsCompliant ) {$AADComputer.IsCompliant } else {''}
                    'IsManaged' = if ($AADComputer.IsManaged ) {$AADComputer.IsManaged } else {''}
                   'LastDirSyncTime' = if ($Null -ne $AADComputer.LastDirSyncTime) {$AADComputer.LastDirSyncTime } else {''}
                    'AADRegisteredOwnerUPN' =  if ($AADComputer.AADRegisteredOwnerUPN ) {$AADComputer.AADRegisteredOwnerUPN } else {''}
                    'AADRegisteredOwnerEnabled' =  if ($AADComputer.AADRegisteredOwnerEnabled ) {$AADComputer.AADRegisteredOwnerEnabled } else {''}
                   'managedDeviceOwnerType' = if ($AADComputer.managedDeviceOwnerType ) {$AADComputer.managedDeviceOwnerType } else {''}
                    'enrolledDateTime' =  if ($AADComputer.enrolledDateTime ) {$AADComputer.enrolledDateTime } else {''}
                   'lastSyncDateTime' = if ($AADComputer.lastSyncDateTime -like '01/01/0001*') {''}  elseif ($AADComputer.lastSyncDateTime ) {$AADComputer.lastSyncDateTime } else {''} 
                    'deviceEnrollmentType' =  if ($AADComputer.deviceEnrollmentType ) {$AADComputer.deviceEnrollmentType } else {''}
                    'isEncrypted' =   if ($AADComputer.isEncrypted ) {$AADComputer.isEncrypted } else {''}
                    'model' =  if ( $AADComputer.model) {$AADComputer.model } else {''}
                    'manufacturer' = if ($Null -ne $AADComputer.manufacturer) {$AADComputer.manufacturer } else {''}
                    'serialNumber' =  if ($AADComputer.serialNumber ) {$AADComputer.serialNumber } else {''}
                    'IntuneManaged' = $IntuneManaged
                    'DateDeleted' = ''
                    'DateDisabled' = if (($AADComputer.AccountEnabled -eq $False ) -and (!$CurrentRecord.DateDisabled)) { get-date } elseif ($AADComputer.AccountEnabled -eq $True) { '' }
                }
    
                $Hostname = $AADComputer.DisplayName 
                #write-verbose "Uploading data for $Hostname"  #Comment out if not required, fills up logs
    
                if ($Null -eq $CurrentRecord) {Add-azTableRow -table $AADComputersTable -partitionKey $partitionKey -rowKey ((New-Guid).Guid) -property $PropertyArray | out-null  }
                else { 
                    $CurrentRecord.Name = $PropertyArray.Name
                    $CurrentRecord.ObjectID = $PropertyArray.ObjectID
                    $CurrentRecord.DeviceID = $PropertyArray.DeviceID
                    $CurrentRecord.AccountEnabled = $PropertyArray.AccountEnabled
                    $CurrentRecord.ApproximateLastLogonTimeStamp = $PropertyArray.ApproximateLastLogonTimeStamp
                    $CurrentRecord.DeviceOSType = $PropertyArray.DeviceOSType
                    $CurrentRecord.DeviceOSVersion = $PropertyArray.DeviceOSVersion 
                    $CurrentRecord.DeviceTrustType = $PropertyArray.DeviceTrustType
                    $CurrentRecord.DirSyncEnabled = $PropertyArray.DirSyncEnabled
                    $CurrentRecord.IsCompliant = $PropertyArray.IsCompliant 
                    $CurrentRecord.IsManaged = $PropertyArray.IsManaged
                    $CurrentRecord.LastDirSyncTime = $PropertyArray.LastDirSyncTime
                    $CurrentRecord.AADRegisteredOwnerUPN = $PropertyArray.AADRegisteredOwnerUPN
                    $CurrentRecord.AADRegisteredOwnerEnabled = $PropertyArray.AADRegisteredOwnerEnabled
                    $CurrentRecord.managedDeviceOwnerType = $PropertyArray.managedDeviceOwnerType
                    $CurrentRecord.enrolledDateTime = $PropertyArray.enrolledDateTime
                    $CurrentRecord.lastSyncDateTime = $PropertyArray.lastSyncDateTime
                    $CurrentRecord.deviceEnrollmentType = $PropertyArray.deviceEnrollmentType
                    $CurrentRecord.isEncrypted = $PropertyArray.isEncrypted
                    $CurrentRecord.model = $PropertyArray.model
                    $CurrentRecord.manufacturer = $PropertyArray.manufacturer
                    $CurrentRecord.serialNumber = $PropertyArray.serialNumber
                    $CurrentRecord.IntuneManaged = $PropertyArray.IntuneManaged
                    $CurrentRecord.DateDeleted = $PropertyArray.DateDeleted
                    if ($PropertyArray.DateDisabled) {$CurrentRecord.DateDisabled = $PropertyArray.DateDisabled}
                    $CurrentRecord | update-aztablerow -table $AADComputersTable  |  out-null    }
                }
            catch {
                write-warning "Error updating computer $Hostname"
                Write-warning "Error: $($_.Exception.Message)"
                } 
            
           } 
        
    
    }
    
    Catch {
    
    Write-warning "Error: $($_.Exception.Message)"
    
    Break
     }
     }
     End {
     If ($?) {
        write-verbose 'Completed updating Azure AD computer info to Azure Table successfully.' 
    }
      }
}
#
function Upload-AADComputerMetrics {

    $PropertyArray = @{
                
        'TotalComputers' = ($AADComputers).count
        'EnabledTrue' = ($AADComputers | where {$_.AccountEnabled -eq $True}).count
        'EnabledFalse' = ($AADComputers | where {$_.AccountEnabled -eq $False}).count
        'DeviceTrustTypeAzureAD' = ($AADComputers | where {$_.DeviceTrustType -like 'AzureAD'}).count
        'DeviceTrustTypeWorkplace' = ($AADComputers | where {$_.DeviceTrustType -like 'Workplace'}).count
        'DeviceTrustTypeServerAD' = ($AADComputers | where {$_.DeviceTrustType -like 'ServerAD'}).count
        'IntuneManagedTrue' = ($AADComputers | where {$_.IntuneManaged -like 'True'}).count
        'IntuneManagedFalse' = ($AADComputers | where {$_.IntuneManaged -like 'False'}).count
        'DeviceEnrollmentTypeWindowsAzureADJoin' = ($AADComputers | where {$_.DeviceEnrollmentType -like 'WindowsAzureADJoin'}).count
        'DeviceEnrollmentTypeWindowsCoManagement' = ($AADComputers | where {$_.DeviceEnrollmentType -like 'WindowsCoManagement'}).count
        'DeviceEnrollmentTypeWindowsAutoEnrollment' = ($AADComputers | where {$_.DeviceEnrollmentType -like 'WindowsAutoEnrollment'}).count
		'DeviceEnrollmentTypeUserEnrollment' = ($AADComputers | where {$_.DeviceEnrollmentType -like 'UserEnrollment'}).count
    }
    
    Add-azTableRow -table $AADComputerMetricsTable -partitionKey $partitionKey -rowKey ((New-Guid).Guid) -property $PropertyArray | out-null
}
#
Function Disable-StaleAADComputer {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DisableDays = '30'

    )

    Begin {

        $DisableTime = (Get-Date).Adddays(-($DisableDays))  

        write-verbose "Azure Active Directory Computer accounts not logged into since $DisableTime will be disabled."

    }

    Process {
    Try {
        foreach ($AADComputer in $AADComputers) {
            
            $ComputerGUID = $AADComputer.ObjectId
            $DeviceID = $AADComputer.DeviceID
            $Name = $AADComputer.DisplayName
            $OwnerUPN = $AADComputer.AADRegisteredOwnerUPN
            $LastLogon = $AADComputer.ApproximateLastLogonTimeStamp
            $TrustType = $AADComputer.devicetrusttype

            if ( $AADComputer.ApproximateLastLogonTimeStamp -lt $DisableTime -and $AADComputer.AccountEnabled -eq $True -and $TrustType -notlike 'ServerAd' ) { 
                    
                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'Owner UPN' = $OwnerUPN
                    'Last Logon' = $LastLogon
                    'Device Trust Type' = $TrustType
                    'Object ID' = $ComputerGUID
                    'Device ID' = $DeviceID
                    }
                $DisabledAADComputerTable.add($Obj)       
                
                $CurrentRecord = get-azTableRow -table $AADComputersTable -columnName "ObjectID" -Value "$ComputerGUID"  -operator Equal

                write-verbose "Disabling AAD computer account for $Name"
                Set-AzureADDevice -objectid $AADComputer.ObjectId -AccountEnabled $False # -whatif not supported on this cmdlet :( Comment the whole line out for dry runs
                
                $CurrentRecord | add-member 'DateDisabled' (get-date) -force 
                $CurrentRecord | update-aztablerow -table $AADComputersTable  |  out-null 
                }
            }
            write-verbose 'Successfully completed disabling AAD computers.'
    } 
    Catch {

        Write-warning -BackgroundColor Red "Error: $($_.Exception.Message)"

        Break
            }
    }
}
#

Function Delete-StaleAADComputer {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $False, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DeleteDays = '60'
    )
    Begin {
        $DeleteTime = (Get-Date).Adddays(-($DeleteDays))  
        write-verbose "Azure Active Directory Computer accounts not logged into since $DeleteTime will be disabled."
    }

    Process {
    Try {
        foreach ($AADComputer in $AADComputers) {
            
            $ComputerGUID = $AADComputer.ObjectId
            $DeviceID = $AADComputer.DeviceID
            $Name = $AADComputer.DisplayName
            $OwnerUPN = $AADComputer.AADRegisteredOwnerUPN
            $LastLogon = $AADComputer.ApproximateLastLogonTimeStamp
            $TrustType = $AADComputer.devicetrusttype

            if ($LastLogon -lt $DeleteTime -and $AADComputer.AccountEnabled -eq $False -and $TrustType -notlike 'ServerAD' ) { 
                  
                $Obj = [pscustomobject] @{
                    'Name' = $Name
                    'Owner UPN' = $OwnerUPN
                    'Last Logon' = $LastLogon
                    'Device Trust Type' = $TrustType
                    'Object ID' = $ComputerGUID
                    'Device ID' = $DeviceID
                    }
                $DeletedAADComputerTable.add($Obj)       
                
                $CurrentRecord = get-azTableRow -table $AADComputersTable -columnName "ObjectID" -Value "$ComputerGUID"  -operator Equal

                write-verbose "Deleting AAD computer account for $Name"
                if ($AADComputer.IntuneManaged -eq $True) {remove-intunemanageddevice -manageddeviceID $AADComputer.deviceID  } #comment out for dry runs
                Remove-AzureADDevice -objectid $AADComputer.ObjectId   # -whatif not supported on this cmdlet :( comment the whole line out for dry runs
                
                $CurrentRecord | add-member 'DateDeleted' (get-date) -force 
                $CurrentRecord | update-aztablerow -table $AADComputersTable  |  out-null 
                }
            }
            write-verbose 'Successfully completed deleting AAD computers.'
    } 
    Catch {

        Write-warning -BackgroundColor Red "Error: $($_.Exception.Message)"

        Break
            }
    }
}

##
function Email-UserComputerReport {
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string]$DestinationEmail
    )
 
Email -AttachSelf -AttachSelfName "User and Computer Report for $TodaysDate" {
    EmailHeader {
        EmailFrom -Address 'YourSender@yourdomain.com'
        EmailTo -Addresses "$DestinationEmail"
        EmailServer -Server 'yourmailserver.domain.com'
        EmailSubject -Subject "User and Computer Report for $TodaysDate"
    }
    EmailBody {
        EmailTextBox -FontFamily 'Calibri' -Size 17 -TextDecoration underline -Color DarkBlue -Alignment center {
            "User and Computer Report for $TodaysDate"
        }
        EmailText -LineBreak
        #EmailText -FontFamily 'Calibri' -Size 20 -color Red -text 'Testing Only - no accounts disabled or deleted (yet)'
        EmailText -LineBreak
        EmailText -FontFamily 'Calibri' -Size 15 -text 'The following user and computer accounts have been disabled or deleted in line with our standard auditing and cleanup process.'
        EmailText -FontFamily 'Calibri' -Size 15 -text 'Accounts disabled today will be deleted in a further 30 days ', 'if they are still not used, or added to a protection group. If you have concerns about one of the accounts in this list, 
        please speak to your local friendly Tech Team support person.' -color Red, Black
        EmailText -LineBreak
        EmailText -FontFamily 'Calibri' -Size 15 -text 'Please review the documentation for this automated process: URL to help doc here...'
        EmailText -LineBreak
            if ($DisabledUserTable -ne $Null) {
            EmailText -FontFamily 'Calibri' -Size 18 -color Blue -text 'The following user accounts have been disabled as they have had no login activity for 30 days.'
            EmailText -LineBreak
            EmailTable -Table $DisabledUserTable
            EmailText -LineBreak
        }
            if ($DeletedUserTable -ne $Null) {
            EmailText -FontFamily 'Calibri' -Size 18 -color Blue -text 'The following user accounts have been deleted as they have had no login activity for 60 days.'
            EmailText -LineBreak
            EmailTable -Table $DeletedUserTable
            EmailText -LineBreak
        }
        if ($DisabledComputerTable -ne $Null) {
            EmailText -FontFamily 'Calibri' -Size 18 -color Blue -text 'The following computer accounts have been disabled as they have had no login activity for 30 days.'
            EmailText -LineBreak
            EmailTable -Table $DisabledComputerTable
            EmailText -LineBreak
        }
        if ($DeletedComputerTable -ne $Null) {
            EmailText -FontFamily 'Calibri' -Size 18 -color Blue -text 'The following computer accounts have been deleted as they have had no login activity for 60 days.'
            EmailText -LineBreak
            EmailTable -Table $DeletedComputerTable
            EmailText -LineBreak
        }
        if ($DisabledAADComputerTable -ne $Null) {
            EmailText -FontFamily 'Calibri' -Size 18 -color Blue -text 'The following Azure AD computer accounts have been disabled as they have had no login activity for 30 days.'
            EmailText -LineBreak
            EmailTable -Table $DisabledAADComputerTable
            EmailText -LineBreak
        }
        if ($DeletedAADComputerTable -ne $Null) {
            EmailText -FontFamily 'Calibri' -Size 18 -color Blue -text 'The following Azure AD computer accounts have been deleted as they have had no login activity for 60 days.'
            EmailText -LineBreak
            EmailTable -Table $DeletedAADComputerTable
            EmailText -LineBreak
        }
        EmailTextBox -FontFamily 'Calibri' -Size 11 -fontstyle italic -Color Blue -Alignment right {
            "$MOTD"
        }
    }
}
}


function Update-ADDeletedUsers {
    <#
    .SYNOPSIS
    Uploads data for recently deleted user accounts to Azure Table Storage. 
    .DESCRIPTION
    Queries AD for recently deleted user objects, looks these objects up in AZ Table based on GUID and if the Table copy is DateDeleted $Null, updates said field. This accommodates
    for user deletions that happen outside this scripted process.  A more streamlined approach to this would be using a SIEM or similar to call a webhook for a Runbook when it detects the event type from AD. 
    .EXAMPLE
    Update-ADDeletedUsers
    .INPUTS
    N/A
    .OUTPUTS
    All data is uploaded directly to the Azure Storage Table defined in the Global Declarations of this script.
    .NOTES
    Its possible a deleted AD account may be detected, that does not exist in the Azure Table. This would happen if the account was both created and deleted between this AD audit process runs. 
    Assuming the process runs daily, accounts detected in this manner probably don't matter. If the process does not run for extended time, its possible we may miss actual/real user accounts. 
#>
    begin {
        $ADDeletedUsers = Get-ADObject -IncludeDeletedObjects -Filter {objectClass -eq "user" -and IsDeleted -eq $True} -Properties displayname, userprincipalname, whencreated,whenchanged, ObjectGUID

    }
    
    process {
        foreach ($user in $ADDeletedUsers) {
            $UserGUID = $User.Objectguid
            $ADDeletedDate = $user.whenchanged
            $UPN = $user.userprincipalname
            $CurrentRecord = get-azTableRow -table $table -columnName "ObjectGUID" -guidValue ([guid]"$UserGUID")  -operator Equal
        
            if ($Null -ne $CurrentRecord) {
                if (($Null -eq $CurrentRecord.DateDeleted) -or ($CurrentRecord.datedeleted -eq '')) {
                    # update the az table here with date deleted record
                    $CurrentRecord | add-member 'DateDeleted' "$ADDeletedDate" -force
                    $CurrentRecord.Enabled = 'False' #Just to assist with reporting
                    $CurrentRecord | update-aztablerow -table $table  |  out-null
                    write-verbose "User $UPN was deleted in AD, have updated the DeleteDate in AZ Table `n"
                }
                elseif ($Null -ne $CurrentRecord.DateDeleted) {
                    write-verbose "User $UPN already listed as deleted in AZ Table `n"
                }
            }
            elseif ($Null -eq $CurrentRecord) {
                write-verbose "User $UPN is deleted in AD but does not exist in AZ Table. This can happen if the user is created and deleted before the AD audit process runs. 
                Assuming this process runs regularly, this probably doesn't matter and the account was likely just a test or mistake. We're not updating the AZ Table here. `n"
            }
        }
    }
    
    end {
        write-verbose "Finished updating AD Deleted Users"
    }
}

function Update-ADDeletedComputers {
    <#
    .SYNOPSIS
    Uploads data for recently deleted computer accounts to Azure Table Storage. 
    .DESCRIPTION
    Queries AD for recently deleted computer objects, looks these objects up in AZ Table based on GUID and if the Table copy is DateDeleted $Null, updates said field. This accommodates
    for computer deletions that happen outside this scripted process.  
    .EXAMPLE
    Update-ADDeletedcomputers
    .INPUTS
    N/A
    .OUTPUTS
    All data is uploaded directly to the Azure Storage Table defined in the Global Declarations of this script.
    .NOTES
    Its possible a deleted AD account may be detected, that does not exist in the Azure Table. This would happen if the account was both created and deleted between AD audit process runs. 
    Assuming the process runs daily, accounts detected in this manner probably don't matter. If the process does not run for extended time, its possible we may miss actual/real computer accounts. 
#>
    begin {
        $ADDeletedComputers = Get-ADObject -IncludeDeletedObjects -Filter {objectClass -eq "computer" -and IsDeleted -eq $True} -properties msDS-LastKnownRDN,objectguid,created,modified

    }
    
    process {
        foreach ($Computer in $ADDeletedComputers) {
            $ComputerGUID = $Computer.Objectguid
            $ADDeletedDate = $Computer.modified
            $Name = $Computer.'msDS-LastKnownRDN'
            $CurrentRecord = get-azTableRow -table $ComputersTable -columnName "ObjectGUID" -guidValue ([guid]"$ComputerGUID")  -operator Equal
        
            if ($Null -ne $CurrentRecord) {
                if (($Null -eq $CurrentRecord.DateDeleted) -or ($CurrentRecord.datedeleted -eq '')) {
                    # update the az table here with date deleted record
                    $CurrentRecord | add-member 'DateDeleted' "$ADDeletedDate" -force
                    $CurrentRecord.Enabled = 'False' #Just to assist with reporting
                    $CurrentRecord | update-aztablerow -table $ComputersTable  |  out-null
                    write-verbose "Computer $Name was deleted in AD, have updated the DeleteDate in AZ Table `n"
                }
                elseif ($Null -ne $CurrentRecord.DateDeleted) {
                    write-verbose "Computer $Name already listed as deleted in AZ Table `n"
                }
            }
            elseif ($Null -eq $CurrentRecord) {
                write-verbose "Computer $Name is deleted in AD but does not exist in AZ Table. This can happen if the computer is created and deleted before the AD audit process runs. 
                Assuming this process runs regularly, this probably doesn't matter and the account was likely just a test or mistake. We're not updating the AZ Table here. `n"
            }
        }
    }
    
    end {
        write-verbose "Finished updating AD Deleted Computers"
    }
}

function Update-AADDeletedComputers {
    <#
    .SYNOPSIS
    Uploads data for recently deleted Azure AD Computer accounts to Azure Table Storage. 
    .DESCRIPTION
     Queries AZ table storage for a list of devices that we think still exist. Cross checks that list against a query of live AAD data to confirm the device does in fact still exist. 
     If it does not exist, updates the DateDeleted field with the date the script ran. If this script does not run daily, the DateDeleted field will not be 100% accurate. 
    .EXAMPLE
    Update-AADDeletedComputers
    .INPUTS
    N/A
    .OUTPUTS
    All data is uploaded directly to the Azure Storage Table defined in the Global Declarations of this script.
    .NOTES
    Azure AD has no recycle bin for deleted computers. This means we have to query table storage for a list of devices we think still exist, then query AD to see if it does actually exist. 
    This is really inefficient and I dont like it. The alternative is Azure Monitor > Alert > Runbook, but that needs more investigation.  See: https://docs.microsoft.com/en-us/azure/automation/automation-create-alert-triggered-runbook
    AAD Recycle Bin for Computers is highly requested since 2017: https://feedback.azure.com/forums/169401-azure-active-directory/suggestions/32127307-recycle-bin-for-deleted-devices 

    #>
    begin {
        $TableRecords = get-aztablerow -table $AADComputersTable -columnname 'DateDeleted' -Value '' -operator equal    }
    
    process {
        foreach ($CurrentRecord in $TableRecords) {
            $ObjectGuid = $CurrentRecord.objectid
            $Name = $CurrentRecord.Name
            $Device = $AADComputers | ? {$_.ObjectID -eq "$ObjectGUID"}
            $Today = get-date
        
            if ($Null -eq $Device) {
                write-verbose "$Name no longer exists in Azure AD. Updating AZ Table. `n"
                $CurrentRecord | add-member 'DateDeleted' "$Today" -force
                $CurrentRecord.AccountEnabled = 'False' #Just to assist with reporting
                $CurrentRecord | update-aztablerow -table $AADComputersTable  |  out-null        
            }
            elseif ($Null -ne $Device) {
                #write-output "Device $Name still exists in AAD, nothing to do here `n"
            }
            }
        }

    end {
        write-verbose "Finished updating AAD Deleted Computers"
    }
}

#endregion

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Script Execution goes here

#Get our user and computer details
write-verbose 'Getting AD Users'
$users = get-aduser -Filter * -resultsetsize $Null -Properties *
write-verbose 'Getting AD Computers'
$Computers = get-adcomputer -Filter * -resultsetsize $Null -Properties *
write-verbose 'Getting AAD Computers'
$AADComputers = get-azureaddevice -all:$True 

# You only need to create a new table column if your adding a new attribute to the data set, normally wont need this. You also need to manually update the parameter set in upload-aduserinfo or other relevant Function. 
 #New-TableColumn -targettablename 'AADComputers1' -NewColumnName 'IntuneManaged'

#Update the user details with the mailbox logon details. This takes a while depending on user/mailbox count as we are making API calls to Exchange Online. 
# You must run this before you can run upload-aduserinfo
write-verbose 'Getting User last mailbox logon'
foreach ($User in $Users) { #Add the users mailbox last logon to the array. 
    $Email = $User.emailaddress
    write-host "$Email" -ForegroundColor green
    if ($Email) {
        $MXLastLogon = Get-MailboxLastLogon -useremail $Email
        $User | add-member 'MailboxLastLogon' "$MXLastLogon" -force
        if ($User.MailboxLastLogon -notlike '') { $User.MailboxLastLogon = [datetime]$User.MailboxLastLogon  } 
    }
    elseif (!($Email)) {
        $MXLastLogon = ''
        $User | add-member 'MailboxLastLogon' "$MXLastLogon" -force
        if ($User.MailboxLastLogon -notlike '') { $User.MailboxLastLogon = [datetime]$User.MailboxLastLogon  } 
    }
}

#Lets clean up first. Updating the deleted/disabled records for any records in Az Table Storage for objects in AD/AAD that were modified outside of this process. 
write-verbose 'Updating table storage for users and devices that have been deleted outside of this process'
Update-ADDeletedUsers

Update-ADDeletedComputers

Update-AADDeletedComputers

#Upload the details to Azure Table
Upload-ADUserInfo -previousdays 10000
Upload-UserMetrics

Disable-StaleADUser -disabledays 30
Delete-StaleADUser -deletedays 60

Upload-ADComputerInfo -previousdays 10000
Upload-ComputerMetrics

Disable-StaleADComputer -disabledays 30
Delete-StaleADComputer -deletedays 60

Upload-AADComputerInfo   # This can take a while depending on number of AAD objects as you make several API calls
Upload-AADComputerMetrics

Disable-StaleAADComputer -disabledays 30
Delete-StaleAADComputer -deletedays 60



#region Sending email notification only if there were actually changes
write-verbose 'Sending email report'
if ($Null -ne $DisabledUserTable -or $Null -ne $DeletedUserTable -or $Null -ne $DisabledComputerTable -or $Null -ne $DeletedComputerTable -or $Null -ne $DisabledAADComputerTable -or $Null -ne $DeletedAADComputerTable) {
    Email-UserComputerReport -destinationemail YourRecipient@YourDomain.com
}
else {
    write-verbose 'No user or computer changes today - not sending any email notifications'
}
#endregion

write-verbose 'We are done here.' 



##############