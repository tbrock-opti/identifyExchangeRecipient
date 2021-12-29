* SYNOPSIS
    * Assists with identifying an exchange recipient and its type
* DESCRIPTION
    * Pulls and parses data from Get-EXORecipient to make it user friendly
* EXAMPLE
    * PS C:\> Start-IdentifyRecipient.ps1 'user@domain.com'
    * Pulls information on the specified email address
* INPUTS
    * Recipient as an alias, email address, or username
* OUTPUTS
    * Recipient
    * Recipient Type
    * Recipient Display Name
    * Recipient Primary Email Address
    * Users with full access (if shared mailbox)
