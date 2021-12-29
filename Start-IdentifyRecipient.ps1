<#
.SYNOPSIS
    Assists with identifying an exchange recipient and its type
.DESCRIPTION
    Pulls and parses data from Get-EXORecipient to make it user friendly
.EXAMPLE
    PS C:\> Start-IdentifyRecipient.ps1 'user@domain.com'
    Pulls information on the specified email address
.INPUTS
    Inputs
        Recipient as an alias, email address, or username
.OUTPUTS
    Output
        Recipient
        Recipient Type
        Recipient Display Name
        Recipient Primary Email Address
        Users with full access (if shared mailbox)
#>


[CmdletBinding()]
param (
    # Recipient to check
    [Parameter(Mandatory)]
    [string]
    $Recipient
)

    
# Check for Connect-ExchangeOnline command to see if module is installed
if (!$(Get-Command 'Connect-ExchangeOnline')) {
        
    # prompt to install
    $message = 'Exchange Online Powershell module not installed, install it? [y/n]' 
    $install = Read-Host -Prompt $message

    # if y, launch an elevated window to install module
    if ($install -like 'y') {
            
        # parameters for launching powershell
        $spArgs = ' -command "& {Install-Module ExchangeOnlineManagement -Force -verbose}" -noexit'
            
        # splat Start-Process parameters
        $spParams = @{
            Verb         = 'runas'
            FilePath     = 'powershell'
            ArgumentList = $spArgs
            Wait         = $true
        }
        Start-Process @spParams

        # module should be installed so we can import it
        Import-Module ExchangeOnlineManagement

    }
}
else {
    Write-Verbose 'ExchangeOnlineManagement is installed.' -Verbose
}

# check if already connected to exchange online
$eoSession = Get-PSSession | Where-Object {
    $_.ComputerName -eq 'outlook.office365.com'
    $_.ConfigurationName -eq 'Microsoft.Exchange' -and 
    $_.Name -like 'ExchangeOnlineInternalSession_*' -and 
    $_.State -eq 'Opened'
}

if (!$eoSession) {

    # provide some direction
    Write-Verbose -Message 'No active session for Exchange online, connecting...' -Verbose
    Write-Verbose -Message 'User Exchange Online admin credential when prompted...' -Verbose

    # pause 1 second so direction is seen
    Start-Sleep -Seconds 1

    # connect exchange online
    Connect-ExchangeOnline -ShowBanner:$false
} else {
    Write-Verbose 'Using existing session for Exchange Online...' -Verbose
}

# get recipient information
$recipientInfo = Get-EXORecipient $Recipient -ErrorAction SilentlyContinue
if (!$recipientInfo) {
    Write-Verbose 'ERROR: Recipient not found. Double check spelling' -Verbose
    Break
}

# create a lamens terms recipient type from the recipientTypeDetails value
switch ($recipientInfo.RecipientTypeDetails) {
    'MailContact' { $recType = 'Contact' }
    'MailUniversalSecurityGroup' { $recType = 'Mail-enabled Security Group' }
    'MailUniversalDistributionGroup' { $recType = 'Distribution Group' }
    'GroupMailbox' { $recType = 'O365/Teams/Yammer Group' }
    'SharedMailbox' { $recType = 'Shared Mailbox' }
    'UserMailbox' { $recType = 'User Mailbox' }
}

# create an object to return results
$result = [PSCustomObject]@{
    Recipient               = $Recipient
    'Recipient Type'        = $recType
    'Display Name'          = $recipientInfo.DisplayName
    'Primary Email Address' = $recipientInfo.PrimarySmtpAddress
}

# if recipient is a shared mailbox
$faPerms = $null
if ($recType -eq 'Shared Mailbox') {

    # pull users with fullAccess that isn't SELF or a deny line
    $faPerms = Get-MailboxPermission -Identity $recipientInfo.identity | Where-Object {
        $_.accessRights -like '*FullAccess*' -and 
        $_.Deny -eq $false -and 
        $_.User -ne 'NT AUTHORITY\SELF'
    }
    
    # add the result as a property to the $result object
    $amParams = @{
        MemberType = 'NoteProperty'
        Name       = 'FullAccess Users'
        Value      = $faPerms.User
    }
    $result | Add-Member @amParams
}

# dump results
$result