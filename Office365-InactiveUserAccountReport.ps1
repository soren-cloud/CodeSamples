# Connect to Exchange Online 
function Connect-ExchangeOnline 
    {
    param ($Credential,$Commands)
    try
        {
        Write-Output "Connecting to Exchange Online"
        Get-PSSession | Remove-PSSession       
        $Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential `
            -Authentication Basic -AllowRedirection
        Import-PSSession -Session $Session -DisableNameChecking:$true -AllowClobber:$true -CommandName $Commands | Out-Null
        }
    catch 
        {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
        }
    Write-Verbose "Successfully connected to Exchange Online"
    }

# Disconnect to Exchange Online 
function Disconnect-ExchangeOnline 
    {
    try
        {
        Write-Output "Disconnecting from Exchange Online"
        Get-PSSession | Remove-PSSession       
        }
    catch 
        {
        Write-Error -Message $_.Exception
        Stop-AutomationScript -Status Failed
        }
    Write-Verbose "Successfully disconnected from Exchange Online"
    }

# !!! Missing part: Create RBAC group "View-Only RBAC", that only grants access to view RBAC groups


$Credential = Get-Credential
Connect-ExchangeOnline -Credential $Credential -Commands "Get-Mailbox","Get-RoleGroup","Get-RoleGroupMember","Search-UnifiedAuditLog"
Connect-MsolService -Credential $Credential

# OFFICE 365
# # Get Users
$AllUsers = Get-MsolUser -All | Select-Object DisplayName,UserPrincipalName,IsLicensed,O365AdminRoleType,O365AdminSubRole,ExORBACMembership,ExORecipientType, `
    @{Name="UserState";Expression={if($_.BlockCredential -eq $false){"Enabled"}else{"Disabled"}}},ExOAuditEnabled,ExOAuditLogAgeLimitInDays,LastLogonTimestamp, `
    LastLogonDaysAgo, LastLogonDisplay, @{Name="UserCreatedTimestamp";Expression={$_.WhenCreated}}, `
    @{Name="UserCreatedDaysAgo";Expression={$(New-TimeSpan -Start $_.WhenCreated).Days}}, ObjectId

# # Get Admin Roles
$AdminAssignments = @()
$AdminRolesExcluded = "Partner Tier1 Support","Partner Tier2 Support","Guest User"
$AdminRoles = Get-MsolRole | Where-Object {$AdminRolesExcluded -notcontains $_.Name}
foreach ($AdminRole in $AdminRoles)
    {
    $AdminRoleMembers = Get-MsolRoleMember -RoleObjectId $AdminRole.ObjectId.Guid
    if ($AdminRoleMembers)
        {
        $AdminAssignments += $AdminRoleMembers | Select-Object @{Name="AdminRoleName";Expression={$AdminRole.Name}},ObjectId
        }
    }

# # Add Admin Roles information to Users
foreach ($AdminAssignment in $AdminAssignments)
    {
    if ($AdminAssignment.AdminRoleName -eq "Company Administrator")
        {
        $O365AdminRoleType = "Global Administrator"
        }
    else
        {
        $O365AdminRoleType = "Customized Administrator"
        $O365AdminSubRole = "$($AdminAssignment.AdminRoleName)`n"
        }
    $($AllUsers | Where-Object {$_.ObjectId -eq $AdminAssignment.ObjectId}).O365AdminRoleType = $O365AdminRoleType
    $($AllUsers | Where-Object {$_.ObjectId -eq $AdminAssignment.ObjectId}).O365AdminSubRole += $O365AdminSubRole
    }
# #  Will remove tailing new-line in O365AdminSubRole property
$AllUsers | Where-Object {$_.O365AdminSubRole} | ForEach-Object {$_.O365AdminSubRole = $_.O365AdminSubRole.trim()}


# EXCHANGE
# # Get Mailboxes and add information to users
$Mailboxes = Get-Mailbox -ResultSize Unlimited


# # RBAC again
$ExORBACGroupAssignments = @()
$ExORBACGroups = Get-RoleGroup -ResultSize Unlimited | Where-Object {$_.Members}
foreach ($ExORBACGroup in $ExORBACGroups)
    {
    $ExORBACGroupName = $ExORBACGroup.Identity
    $ExORBACGroupMembers = Get-RoleGroupMember -Identity $ExORBACGroupName -ResultSize Unlimited | Where-Object {$_.RecipientType -eq "User"}
    if ($ExORBACGroupMembers)
        {
        $ExORBACGroupAssignments += $ExORBACGroupMembers | Select-Object @{Name="ExORoleName";Expression={$ExORBACGroupName}}, `
            @{Name="ObjectId";Expression={$_.ExternalDirectoryObjectId}}
        }
    }

# # Add Mailbox information and RBAC membership information to users in list
Foreach ($User in $AllUsers)
    {
    # Mailbox info
    $Mailbox = $Mailboxes | Where-Object {$_.UserPrincipalName -eq $User.UserPrincipalName}
    $User.ExORecipientType = $Mailbox.RecipientTypeDetails
    $User.ExOAuditEnabled = $Mailbox.AuditEnabled
    $ExOAuditLogAgeLimit = $Mailbox.AuditLogAgeLimit
    if ($ExOAuditLogAgeLimit)
        {
        $User.ExOAuditLogAgeLimitInDays = $($Mailbox.AuditLogAgeLimit.Split("."))[0]
        }
    $User.ExORBACMembership = $(($ExORBACGroupAssignments | Where-Object {$_.ObjectId -eq $User.ObjectId}).ExORoleName) -join "`n"
    }



# # Add Last logged in timestamp
# # # Maybe another way to loop following section? --> https://blogs.msdn.microsoft.com/tehnoonr/2018/01/26/retrieving-office-365-audit-data-using-powershell/
Foreach ($User in $AllUsers)
    {
    $DaysAgoSpan = $User.UserCreatedDaysAgo
    if ($DaysAgoSpan -gt 90)
        {
        $DaysAgoSpan = 90
        }
    $AuditLog = Search-UnifiedAuditLog -UserIds $User.UserPrincipalName -StartDate $(Get-Date).adddays(-$DaysAgoSpan) `
        -EndDate $(Get-Date) -ResultSize 5000 -Operations UserLoggedIn| Sort-Object CreationDate | Select-Object -Last 1
    $LastLogonTimestamp = $AuditLog.CreationDate
    if ($LastLogonTimestamp)
        {
        $LastLogonDaysAgo = $(New-TimeSpan -Start $LastLogonTimestamp).Days
        $User.LastLogonDaysAgo = $LastLogonDaysAgo
        }
    else
        {
        $LastLogonDaysAgo = "Unknown"
        }
    # Add Days Ago Display Property
    if ($LastLogonDaysAgo -ge 2)
        {
        $LastLogonDisplay = "$LastLogonDaysAgo days ago"
        }
    elseif ($LastLogonDaysAgo -le 1)
        {
        $LastLogonDisplay = "One day ago or less"
        }
    elseif ($LastLogonDaysAgo -eq "Unknown" -and $($User.ExOAuditLogAgeLimitInDays))
        {
        $LastLogonDisplay = "$($User.ExOAuditLogAgeLimitInDays), more or never"
        }
    else
        {
        $LastLogonDisplay = "More than 90 days ago or never logged in"
        }
    # Write to user item
    $User.LastLogonTimestamp = $LastLogonTimestamp
    $User.LastLogonDaysAgo = $LastLogonDaysAgo
    $User.LastLogonDisplay = $LastLogonDisplay
    }

# BUILD REPORT
# # Ref: https://azurefieldnotesblog.blob.core.windows.net/wp-content/2017/06/Help-ReportHTML2.html
# # Ref: https://www.powershellgallery.com/packages/ReportHTML
Import-Module ReportHtml
$ReportOutputPath = "C:\temp\testreport.html"

$ReportData_All = $AllUsers `
    | Select-Object @{Name="Display Name";Expression={$_.DisplayName}},@{Name="User Principal Name";Expression={$_.UserPrincipalName}}, `
    @{Name="Last Sign-in";Expression={$_.LastLogonDisplay}},@{Name="User State";Expression={$_.UserState}},@{Name="O365 Admin Role Type";Expression={$_.O365AdminRoleType}}, `
    @{Name="O365 Admin Sub-role";Expression={$_.O365AdminSubRole}},@{Name="Exchange Online RBAC Membership";Expression={$_.ExORBACMembership}},*

$ReportData_All_InScope = $ReportData_All | Where-Object {(($_.ExORecipientType -eq "UserMailbox") -or (!$_.ExORecipientType)) -and  ($_.UserState -eq "Enabled")}

$ReportData_O365Admins = $ReportData_All_InScope | Where-Object {$_.O365AdminRoleType}
$ReportData_ExOAdmins = $ReportData_All_InScope | Where-Object {!$_.O365AdminRoleType -and $_.ExORBACMembership}
$AllUsers_RegularUsers = $ReportData_All_InScope | Where-Object {!$_.O365AdminRoleType -and !$_.ExORBACMembership}


$ReportName = "Office 365 - Inactive User Account Report"
$rpt = @()
$rpt += Get-HTMLOpenPage -TitleText $ReportName

$rpt += Get-HtmlContentOpen -HeaderText "All Users"
$rpt += Get-HtmlContentTable $ReportData_All
$rpt += Get-HtmlContentClose
$rpt += Get-HTMLClosePage -FooterText "Enter Your Custom Text Here" 

$rpt | set-content -path $ReportOutputPath 
Set-Content -Value $rpt -path $ReportOutputPath 
Invoke-item $ReportOutputPath