# HelloID-Task-SA-Target-ExchangeOnPremises-MailboxSetPrimaryAddress
####################################################################
# Form mapping
$formObject = @{
    MailboxIdentity = $form.MailboxIdentity
    DisplayName     = $form.MailBoxDisplayName
    isPrimary       = $form.isPrimary
    newPrimaryMail  = $form.NewPrimaryMail
}

[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnPremises action: [MailboxSetPrimaryAddress] for: [$($formObject.DisplayName)]"
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Set-Mailbox'
    $IsConnected = $true

    if ($formObject.isPrimary -ne $true) {
        $currentMailbox = Get-Mailbox -Identity $formObject.MailboxIdentity
        $list = [System.Collections.ArrayList]::new()
        foreach ($address in $currentMailbox.EmailAddresses) {
            $prefix = $address.Split(':')[0]
            $mail = $address.Split(':')[1]
            if ($mail.ToLower() -eq $formObject.newPrimaryMail.ToLower()) {
                $address = 'SMTP:' + $mail
            } else {
                $address = $prefix.ToLower() + ':' + $mail
            }
            $null = $list.Add($address)
        }
        $paramsSetMailbox = @{
            Identity                  = $formObject.MailboxIdentity
            EmailAddresses            = $list
            EmailAddressPolicyEnabled = $false
        }
        $null = Set-Mailbox @paramsSetMailbox -ErrorAction Stop
    }

    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.DisplayName
        Message           = "ExchangeOnPremises action: [MailboxSetPrimaryAddress] [$($formObject.newPrimaryMail)] for: [$($formObject.DisplayName)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnPremises action: [MailboxSetPrimaryAddress] [$($formObject.newPrimaryMail)] for: [$($formObject.DisplayName)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.DisplayName
        Message           = "Could not execute ExchangeOnPremises action: [MailboxSetPrimaryAddress][$($formObject.newPrimaryMail) for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [MailboxSetPrimaryAddress][$($formObject.newPrimaryMail) for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
####################################################################
