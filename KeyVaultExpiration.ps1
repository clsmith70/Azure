<#
    .SYNOPSIS
        Connects to an Azure Key Vault and retrieves the list of keys, secrets, and certificates and sends an expiration email.

    .DESCRIPTION
        Using the provided parameter data, a connection to Azure, and subsequently a key vault, are made and the list of keys,
        secrects, and certificates are gathered.  The records are then examined for their expiration date and processed according
        to the selected Range.  If a list of items meets the criteria, it is formatted into an HTML table and emailed to the address
        specified.  If any errors occur, they are sent to the AdminEmail specified.

    .PARAMETER PSAutomationCredential
        The Service Principal credential asset to get and use for the connection to Azure

    .PARAMETER TenantId
        The Azure Active Directory TenantId to use to establish the connection

    .PARAMETER TargetCloud
        The Azure Cloud environment to connect to.  The valid values are AzureCloud, AzureChinaCloud, AzureUSGovernment, and AzureGermanCloud.
        The default value is AzureCloud.

    .PARAMETER KeyVaultName
        The name of the Azure Key Vault to read data from, the Service Principal must have access to the vault

    .PARAMETER Range
        The time period to use in deciding which data to return.  The valid values are:
            0 - return only expired items
            1 - return all expired and expiring items out to 90 days (default)
            30 - return all items expiring from today out 30 days
            60 - return all items expiring between 30 and 60 days from today
            90 - return all items expiring between 60 and 90 days from today

    .PARAMETER SendTo
        A valid email address to send the expiriation report to

    .PARAMETER AdminEmail
        A valid email address to send any error reports to

    .PARAMETER FromAddress
        A valid email address to send from, it should match the primary or any alias addresses of the Office 365 account used.

    .PARAMETER SmtpCredential
        The credential asset to get and use to send email messages via Office 365.
#>

Param (
    [Parameter(Mandatory=$true)]
    [string] $PSAutomationCredential,

    [Parameter(Mandatory=$true)]
    [string] $TenantId,

    [Parameter(Mandatory=$false)]
    [ValidateSet('AzureCloud','AzureChinaCloud','AzureUSGovernment','AzureGermanCloud')]
    [string] $Environment = 'AzureCloud',

    [Parameter(Mandatory=$true)]
    [string] $KeyVaultName,

    [Parameter(Mandatory=$false)]
    [ValidateSet(0, 1, 30, 60, 90)]
    [int] $Range = 1,

    [Parameter(Mandatory=$true)]
    [ValidatePattern("^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$")]
    [string] $SendTo,

    [Parameter(Mandatory=$true)]
    [ValidatePattern("^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$")]
    [string] $AdminEmail,

    [Parameter(Mandatory=$true)]
    [ValidatePattern("^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$")]
    [string] $FromAddress,

    [Parameter(Mandatory=$true)]
    [string] $SmtpAutomationCredential
)

$Office365SmtpData = @{
    SmtpServer = "smtp.office365.com"
    Port = 587
    UseSsl = $true
    Subject = "Azure Key Vault Expirations"
    From = $FromAddress
    BodyAsHtml = $true
}

$vaultItemNames = @{
    PSKeyVaultKeyIdentityItem = "Key"
    PSKeyVaultSecretIdentityItem = "Secret"
    PSKeyVaultCertificateIdentityItem = "Certificate"
}

$emailBody = "<html>`n<head>`n`t<meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii"">`n`t<style>`n`t`tbody { text-align: center; }`n"
$emailBody += "`t`ttable { border-collapse: collapse; margin: auto; }`n`t`ttable, th, tr, td { border: 1px solid black; }`n"
$emailBody += "`t`tth, td { padding: 15px; text-align: left; }`n`t`tth { background-color: #63666A; color: #fff; }`n"
$emailBody += "`t`ttr:nth-child(even) { background-color: #72C8CE; }`n`t`t.expired { color: red; }`n`t</style>`n</head>`n`n"
$emailBody += "<body>`n`t<h2>Azure KeyVault Expiration Report</h2>`n`t<h3>Key Vault: $KeyVaultName</h3>`n`t`t<table>`n"

# loop through the list presented and return object data
function ProcessList {
    <#
        .SYNOPSIS
            Takes a list of vault items and returns the items that match the specified range

        .DESCRIPTION
            Takes a list of vault items gathered and loops through them to find the items that match
            the range specified.  Any items that match are then returned in an object list.

        .PARAMETER List
            The list of vault items to process

        .PARAMETER Range
            The date range to process this list by.  Valid values are:
                0 for expired items
                30 for items expiring within the next 30 days
                60 for items expiring within the next 30 to 60 days
                90 for items expiring within th enext 60 to 90 days
    #>
    Param (
        [Parameter(Mandatory=$true)]
        [Object] $List,

        [Parameter(Mandatory=$true)]
        [ValidateSet(0, 30, 60, 90)]
        [int] $Range
    )
    Begin {
    }

    Process {
        $outputList = switch ($range) {
            # expired items
            0 {
                $list | ForEach-Object {
                    if ($_.Expires -le (Get-Date)) {
                        $type = $_.GetType()
                        [pscustomobject] @{
                            Name = $_.Name
                            Type = $vaultItemNames[$type.Name]
                            ExpirationRange = 'Expired'
                            Expires = $_.Expires
                        }
                    }
                }
            }
            # today to 30 days from now
            30 {
                $list | ForEach-Object {
                    if ($_.Expires -le (Get-Date).AddDays(30) -and $_.Expires -ge (Get-Date)) { 
                        $type = $_.GetType()
                        [pscustomobject] @{
                            Name = $_.Name
                            Type = $vaultItemNames[$type.Name]
                            ExpirationRange = '30 Days'
                            Expires = $_.Expires
                        }
                     }
                }
            }
            # 30 to 60 days from now
            60 {
                $list | ForEach-Object {
                    if ($_.Expires -le (Get-Date).AddDays(60) -and $_.Expires -ge (Get-Date).AddDays(30)) {
                        $type = $_.GetType()
                        [pscustomobject] @{
                            Name = $_.Name
                            Type = $vaultItemNames[$type.Name]
                            ExpirationRange = '60 Days'
                            Expires = $_.Expires
                        }
                    }
                }
            }
            # 60 to 90 days from now
            90 {
                $list | ForEach-Object {
                    if ($_.Expires -le (Get-Date).AddDays(90) -and $_.Expires -ge (Get-Date).AddDays(60)) {
                        $type = $_.GetType()
                        [pscustomobject] @{
                            Name = $_.Name
                            Type = $vaultItemNames[$type.Name]
                            ExpirationRange = '90 Days'
                            Expires = $_.Expires
                        }
                    }
                }
            }
        }
    }
    End {
        # return the generated list
        return $outputList
    }
}

try {
    # create arrays to collect information
    $expired = @()
    $outList = @()

    # Get stored credentials for processing
    $vaultCredential = Get-AutomationPSCredential -Name $PSAutomationCredential
    $smtpCredential = Get-AutomationPSCredential -Name $SmtpAutomationCredential

    # Log into the AzAccount
    Connect-AzAccount -ServicePrincipal -Tenant $TenantId -Credential $vaultCredential -Environment $Environment

    # Get the keys in the vault
    $keyList = Get-AzKeyVaultKey -VaultName $KeyVaultName
    # Get the secrets in the vault
    $secretList = Get-AzKeyVaultSecret -VaultName $KeyVaultName
    # Get the certificates in the vault
    $certificateList = Get-AzKeyVaultCertificate -VaultName $KeyVaultName

    # if the requested Range is not for expired items only, get that list
    if ($Range -ne 0) {
        # loop through all 3 lists to get expired items
        $expired += $keyList,$secretList,$certificateList | ForEach-Object {
            # if the list isn't empty, process it
            if ($_) {
                ProcessList -List $_ -Range 0
            } # end If
        } # end foreach list
    } # end if Range -ne 0

    # if Range is 1, process all expiration timeframes
    if ($Range -eq 1) {
        # store the results as they are gathered from each of the lists
        $outList += $keyList,$secretList,$certificateList | ForEach-Object {
            # if the list isn't empty...
            if ($_) {
                # save it for the next loop and process it for all ranges
                $list = $_
                30,60,90 | ForEach-Object {
                    ProcessList -List $list -Range $_
                } # end foreach range
            } # end if
        } # end foreach list
    # otherwise, process for the requested range
    } # end Range -eq 1 if 
    else {
        # store the results as they are gathered from each of the lists
        $outList += $keyList,$secretList,$certificateList | ForEach-Object {
            # if the list isn't empty, process it
            if ($_) {
                ProcessList -List $_ -Range $Range
            } # end if
        } # end foreach list
    } # end Range else

    # if the expired list is not empty, add it to the outList
    if ($expired.Count -gt 0) { $outList += $expired }

    # if the generated list is not empty
    if ($outList.Count -gt 0) {
        # if there are expired items, add them to the list and sort the new list by expiration date
        $outList = $outList | Sort-Object Expires
        # set up the table header row
        $table = "`t`t`t<tr><th>Name</th><th>Type</th><th>ExpirationRange</th><th>Expires</th></tr>`n"
        # loop through each object and return a table row, adding the expired class to expired items
        $outList | ForEach-Object {
            if ($_.Expires -lt (Get-Date)) {
                $table += "`t`t`t<tr class=""expired""><td>$($_.Name)</td><td>$($_.Type)</td><td>$($_.ExpirationRange)</td><td>$($_.Expires)</td></tr>`n"
            } else {
                $table += "`t`t`t<tr><td>$($_.Name)</td><td>$($_.Type)</td><td>$($_.ExpirationRange)</td><td>$($_.Expires)</td></tr>`n"
            }
        }
        # finish the table, body and html
        $table += "`t`t</table>`n`t</body>`n</html>"
    }

    # if there are results to send
    if ($table) {
        $emailBody += $table
        Send-MailMessage @Office365SmtpData -To $SendTo -Credential $smtpCredential -Body $emailBody
    } else {
        Send-MailMessage @Office365SmtpData -To $SendTo -Credential $smtpCredential -Body "<html><body><h3>No expiring items on record</h3></body></html>"
    }
} # end try
catch {
    # Send the error message to someone
    Send-MailMessage @Office365SmtpData -To $AdminEmail -Credential $smtpCredential -Body $($error[0].Exception)
} # end catch