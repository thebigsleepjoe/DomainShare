<#
@thebigsleepjoe on GitHub

Iterates through all mailboxes in EOM and shares each user's calendars matching the below
defined domain with all other users in the same domain. This provides access level Editor.
#>

[CmdletBinding()]
param (
    [string]$targetDomain = "domain.com",
    [string]$adminUpn = "admin@domain.com"
)

# prevent errs from leading @
if ($targetDomain.StartsWith("@")) {
    $targetDomain = $targetDomain.Substring(1)
}

if ("domain.com" -eq $targetDomain -or "admin@domain.com" -eq $adminUpn) {
    Write-Host "[ERR] Usage: .\script -targetDomain <domain.com> -adminUpn <admin@domain.com>"
    return
}

$domain = "@$targetDomain"

Import-Module ExchangeOnlineManagement

Write-Host "Connecting to EOM. Check for any SSO pop-ups!!!"
Connect-ExchangeOnline -UserPrincipalName "$adminUpn" -ShowProgress $true
Write-Host "Finished connecting to EOM"

Write-Host "Grabbing Mailboxes"
$mailboxes = Get-Mailbox -ResultSize Unlimited

# filter by domain
$mailboxes = $mailboxes | Where-Object { $_.PrimarySmtpAddress -like "*$domain" }

# print the output
Write-Host $mailboxes | Format-Table -AutoSize

function Set-ShareCalendarPerms {
    param (
        [string]$calendarOwnerAddress,
        [string]$accessorAddress
    )
    $perms = Get-MailboxFolderPermission -Identity "${calendarOwnerAddress}:\Calendar" -User $accessorAddress -ErrorAction SilentlyContinue

    # check if errored (errors if no access exists)
    if ($null -eq $perms) {
        Add-MailboxFolderPermission -Identity "${calendarOwnerAddress}:\Calendar" -User $accessorAddress -AccessRights Editor -ErrorAction SilentlyContinue
        Write-Host "[ADD] calendar access for $accessorAddress to $calendarOwnerAddress"
        return # done here
    }

    # if access already exists but insufficient
    if ("Editor" -ne $perms.AccessRights) {
        Set-MailboxFolderPermission -Identity "${calendarOwnerAddress}:\Calendar" -User $accessorAddress -AccessRights Editor -ErrorAction SilentlyContinue
        Write-Host "[UPDATE] calendar access for $accessorAddress to $calendarOwnerAddress"
        return
    }

    Write-Host "[SKIP] calendar access for $accessorAddress to $calendarOwnerAddress (user has access already)"
}

function Set-AllShareCalendarPerms {
    param (
        [string]$emailAddress,
        $mailboxes
    )

    foreach ($mb in $mailboxes) {
        $addr = $mb.PrimarySmtpAddress
        if ($addr -eq $emailAddress) { continue }
        Set-ShareCalendarPerms -calendarOwnerAddress $addr -accessorAddress $emailAddress
    }
}

foreach ($mailbox in $mailboxes) {
    $address = $mailbox.PrimarySmtpAddress
    Set-AllShareCalendarPerms -emailAddress $address -mailboxes $mailboxes
}