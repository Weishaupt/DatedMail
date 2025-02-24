#region Variables
[string] $DefaultConfigurationFilePath = Join-Path -Path $env:HOME -ChildPath ".config" -AdditionalChildPath "DatedMail","DatedMailConfig.json"

#endregion Variables

#region Exported Functions

<#
.SYNOPSIS
Initializes a new DatedMail configuration.

.DESCRIPTION
Each DatedMail configuration must contain the path the the Sieve filter script
and the e-mail address where mails for unexpired addresses are forwarded to.

.PARAMETER ConfigurationFilePath
The configuration path for the configuration. By default a configuration is written
to the .config folder in the users home directory at ~/.config/DatedMail/DatedMailConfig.json

.PARAMETER SieveFilterPath
The path to the Sieve filter script which is active for the temp mailbox that should be used

.PARAMETER ForwardingEmailAddress
The fully qualified mail address where mails should be forwarded to.

.PARAMETER MailAddressPrefix
The plus-adressing prefix for the mailbox. This must end with a plus (+) sign
to allow for correct plus addressing. By default 'temp+' is used.

.PARAMETER MailAddressDomain
The mail domain used for all generated mail addresses. E.g. example.org

.EXAMPLE
PS> Initialize-Configuration -SieveFilterPath /home/marvin/users/temp/sieve/active.filter -ForwardingEmailAddress marvin@example.org

This example creates a new configuration with the default configuration path.

#>
function Initialize-DatedMailConfiguration {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([System.Void])]
    param(
        [Parameter(Mandatory=$false)]
        [string] $ConfigurationFilePath = $DefaultConfigurationFilePath,
        [Parameter(Mandatory=$true)]
        [string] $SieveFilterPath,
        [Parameter(Mandatory=$true)]
        [string] $ForwardingEmailAddress,
        [Parameter(Mandatory=$false)]
        [string] $MailAddressPrefix = "temp+",
        [Parameter(Mandatory=$true)]
        [string] $MailAddressDomain
    )

    [PSCustomObject]$config = @{
        Addresses = @();
        MailPrefix = $MailAddressPrefix;
        MailDomain = $MailAddressDomain;
        ForwardingEmailAddress = $ForwardingEmailAddress;
        SieveFilterPath = $SieveFilterPath;
    }

    Test-Configuration -Configuration $config

    Export-Configuration -Configuration $config -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference
}

<#
.SYNOPSIS
Adds a new expiring e-mail address to the database

.PARAMETER ValidDays
Specifies how many days the new generated address should be valid for

.PARAMETER ValidUntil
Specifies the exact date where the generated address should expire

.PARAMETER ValidTimeSpan
Specifies an exact timespan from the current date where the address should expire

.PARAMETER ExportFilePath
Specifies the path to which to export the new email address. The export
will be a text file with only one line consisting of the generated mail address

.PARAMETER ConfigurationFilePath
The configuration path for the configuration. By default a configuration is written
to the .config folder in the users home directory at ~/.config/DatedMail/DatedMailConfig.json

.PARAMETER ReturnMailAddress
Specifies whether the mail address should be returned on STDOUT instead of being written to file

.EXAMPLE
PS> New-DatedMailAddress -ValidDays 7 -ExportFilePath $env:Home/DatedMail

Creates a new mail address valid for exactly 7 days and exports the address to the file
DatedMail in the users home directory.

#>
function New-DatedMailAddress {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([string],[System.Void])]
    param (
        [Parameter(Mandatory=$true, ParameterSetName="Days")]
        [ValidateRange(1,[Int64]::MaxValue)]
        [Int64] $ValidDays,
        [Parameter(Mandatory=$true, ParameterSetName="Date")]
        [datetime] $ValidUntil,
        [Parameter(Mandatory=$true, ParameterSetName="Range")]
        [timespan] $ValidTimeSpan,
        [Parameter(Mandatory=$false)]
        [string] $ExportFilePath = "",
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.SwitchParameter]$ReturnMailAddress,
        [Parameter(Mandatory=$false)]
        [string] $ConfigurationFilePath = $DefaultConfigurationFilePath
    )

    if(![String]::IsNullOrWhiteSpace($ExportFilePath) -and (Test-Path -Path $ExportFilePath -IsValid -PathType Leaf) -eq $false) {
        $ex = New-Object -TypeName System.IO.FileNotFoundException -ArgumentList "The path $ExportFilePath is invalid. Please provide a valid path and try again."
        throw($ex)
    }

    #Get Config
    $config = Import-Configuration -ConfigurationFilePath $ConfigurationFilePath

    $chars = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()

    # Initialize an empty string
    [string] $datedAddress = $config.MailPrefix

    # Generate a random string of length 15
    for ($i = 0; $i -lt 15; $i++) {
        $char = $chars | Get-Random
        $datedAddress += $char
    }

    [datetime]$ExpiryDate = [datetime]::Now
    if($null -ne $ValidDays -and $ValidDays -gt 0) {
        $ExpiryDate = $ExpiryDate.AddDays($ValidDays)
    } elseif ($null -ne $ValidUntil) {
        $ExpiryDate = $ValidUntil
    } elseif ($null -ne $ValidTimeSpan) {
        $ExpiryDate = $ExpiryDate.Add($ValidTimeSpan)
    } else {
        $ex = New-Object -TypeName System.ApplicationException -ArgumentList "Unable to calculate the the expiry date."
        throw($ex)
    }

    if($ExpiryDate -le [datetime]::Now) {
        $ex = New-Object -TypeName System.ApplicationException -ArgumentList "The given expiry date is in the past. Validate your input and try again."
        throw($ex)
    }

    $DatedMail = [PSCustomObject]@{
        Address = "$datedAddress@$($config.MailDomain)";
        ExpiresOn = $ExpiryDate.ToString("s")
    }

    Write-Verbose "Generated address '$($DatedMail.Address)' will expire on $($DatedMail.ExpiresOn)"

    [object[]] $config.Addresses += $DatedMail

    Write-Debug "Adding generated address to existing configuration"
    Export-Configuration -Configuration $config -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference

    Write-Debug "Updating the existing configuration"
    Update-DatedMailAddress -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference -ForceSieveFilterUpdate

    if(![String]::IsNullOrWhiteSpace($ExportFilePath)) {
        Write-Debug "Exporting email address to $ExportFilePath"
        try {
            $DatedMail.Address | Out-File -FilePath $ExportFilePath -WhatIf:$WhatIfPreference -NoNewline -ErrorAction Stop
        }
        catch {
            $ex = New-Object -TypeName System.ApplicationException -ArgumentList "Unable to export address to the given file path $ExportFilePath"
            throw($ex)
        }
    }

    if($ReturnMailAddress -eq $true) {
        return $DatedMail.Address
    } else {
        return
    }
}

<#
.SYNOPSIS
Updates the list of allowed mail addresses based on the current date.

.DESCRIPTION
This function should be called periodically to update the list of allowed
email addresses. Whenever a new dated mail address is created, this function
is automatically called.

This function might update two files:
  1. The SieveFilter file
  2. The config database if expired entries are found

.PARAMETER ConfigurationFilePath
The configuration path for the configuration. By default a configuration is written
to the .config folder in the users home directory at ~/.config/DatedMail/DatedMailConfig.json

.PARAMETER ForceSieveFilterUpdate
Force updating the Sieve filter file, even if no newly expired mail addresses were identified.

.EXAMPLE
PS> Update-DatedMailAddress

Updates the list of allowed mail addresses in the database using the default configuration.

#>
function Update-DatedMailAddress {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([System.Void])]
    param(
        [Parameter(Mandatory=$false)]
        [string] $ConfigurationFilePath = $DefaultConfigurationFilePath,
        [Parameter()]
        [System.Management.Automation.SwitchParameter]$ForceSieveFilterUpdate
    )

    $config = Import-Configuration -ConfigurationFilePath $ConfigurationFilePath
    $DateNow = [datetime]::Now

    Write-Verbose "Removing entries older than $DateNow from the database"

    [bool] $UpdateConfigDatabase = $false

    [PSCustomObject[]]$addresses = foreach($entry in $config.Addresses) {
        $expiry = $entry.ExpiresOn
        if($expiry -gt $DateNow) {
            Write-Debug "Entry with address '$($entry.Address)' will expire at $($entry.ExpiresOn). Keeping."
            $entry
        } else {
            $UpdateConfigDatabase = $true
            Write-Debug "Entry with address '$($entry.Address)' expired on $($entry.ExpiresOn). Removing."
            #Don't return entry as it should be deleted from config db
        }

    }

    if($UpdateConfigDatabase) {
        [int]$expiredAddresses = $config.Addresses.Count - $addresses.Count
        Write-Verbose "$($expiredAddresses) addresses expired since the last execution. Updating config database and Sieve filter."
        $config.Addresses = $addresses
        Update-DatedMailSieveFilter -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference
        Export-Configuration -Configuration $config -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference
    } elseif($ForceSieveFilterUpdate) {
        Write-Verbose "No addresses expired since the last execution. Updating the sieve filter anyways as -ForceSieveFilterUpdate was specified"
        Update-DatedMailSieveFilter -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference
    } else {
        Write-Verbose "No addresses expired. No update required."
    }
}

#endregion Exported Functions

#region Internal Functions
function Update-DatedMailSieveFilter {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$false)]
        [string] $ConfigurationFilePath = $DefaultConfigurationFilePath
    )
    $config = Import-Configuration -ConfigurationFilePath $ConfigurationFilePath

    [string] $header = @"
require ["reject","envelope"];

# This is DatedMail PowerShell Script
# Please don't change anything here.
# Updated: $(Get-Date)


"@

    [string] $footer = @"
reject "550 Invalid or expired recipient address";
stop;
"@

    [string] $template = @"
# Expiry: {1}
if envelope :is "to" "{0}"
{{
  redirect "{2}";
  stop;
}}


"@

    [System.IO.StringWriter] $writer = New-Object -TypeName "System.IO.StringWriter"

    $writer.Write($header)
    foreach($AddressInfo in $config.Addresses) {
        $writer.Write($template -f @($AddressInfo.Address, $AddressInfo.ExpiresOn, $config.ForwardingEmailAddress))
    }
    $writer.Write($footer)

    # Check if path to Sieve file exists, and create it if neccesary
    $SieveContainerPath = Split-Path -Path $config.SieveFilterPath -Parent
    if((Test-Path -Path $SieveContainerPath -PathType Container) -eq $false) {
        New-Item $SieveContainerPath -ItemType Directory -Force -WhatIf:$WhatIfPreference | Out-Null
    }

    $writer.ToString() | Out-File -FilePath $config.SieveFilterPath -WhatIf:$WhatIfPreference
}

function Import-Configuration {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true)]
        [string] $ConfigurationFilePath
    )
    $configFileExists = Test-Path -Path $ConfigurationFilePath -PathType Leaf
    if($configFileExists -eq $true) {
        [PSCustomObject] $config = Get-Content -Path $ConfigurationFilePath | ConvertFrom-Json
    } else {
        $ex = New-Object -TypeName System.ApplicationException -ArgumentList "No configuration file found at the given location $($ConfigurationFilePath). Please run Initialize-Configuration or specify the correct location for the configuration file."
        throw($ex)
    }

    Test-Configuration -Configuration $config

    return $config
}

function Test-Configuration {
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject] $Configuration
    )
    # Prefix
    if($null -eq $Configuration.MailPrefix -or [String]::IsNullOrWhiteSpace($Configuration.MailPrefix)) {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Missing MailPrefix attribute in configuration"
        throw($ex)
    }
    if($Configuration.MailPrefix -notlike "*+") {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Mail Address Prefix must end with '+' to allow for Plus-Adressing"
        throw($ex)
    }

    # Domain
    if($null -eq $Configuration.MailDomain -or [String]::IsNullOrWhiteSpace($Configuration.MailDomain)) {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Missing MailDomain attribute in configuration"
        throw($ex)
    }
    if($Configuration.MailDomain -notlike "*.*") {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Malformed MailDomain: $($Configuration.MailDomain)"
        throw($ex)
    }

    # Sieve Filter Path
    if($null -eq $Configuration.SieveFilterPath -or [String]::IsNullOrWhiteSpace($Configuration.SieveFilterPath)) {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Missing SieveFilterPath attribute in configuration"
        throw($ex)
    }
    if ((Test-Path $Configuration.SieveFilterPath -PathType Leaf -IsValid) -eq $false) {
        $ex = New-Object -TypeName System.IO.IOException -ArgumentList "Unable to validate configuration: Path to Sieve filter file could not be validated. Please check and try again."
        throw($ex)
    }

    # Forwarding Mail Address
    if($null -eq $Configuration.ForwardingEmailAddress -or [String]::IsNullOrWhiteSpace($Configuration.ForwardingEmailAddress)) {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Missing ForwardingEmailAddress attribute in configuration"
        throw($ex)
    }
    try {
        New-Object -TypeName "MailAddress" -ArgumentList @($Configuration.ForwardingEmailAddress) | Out-Null
    } catch [FormatException] {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Invalid forwarding address given: $($Configuration.ForwardingEmailAddress)"
        throw($ex)
    }

    #Addresses
    if($null -eq $Configuration.Addresses) {
        $ex = New-Object -TypeName System.ArgumentException -ArgumentList "Unable to validate configuration: Missing Addresses attribute in configuration"
        throw($ex)
    }
}

function Export-Configuration {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([void])]
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject] $Configuration,
        [Parameter(Mandatory=$true)]
        [string] $ConfigurationFilePath
    )

    # Check if path to config file exists, and create it if neccesary
    $ConfigurationContainerPath = Split-Path $ConfigurationFilePath -Parent
    if((Test-Path $ConfigurationContainerPath -PathType Container) -eq $false) {
        New-Item -Path $ConfigurationContainerPath -ItemType Directory -Force -WhatIf:$WhatIfPreference | Out-Null
    }

    $Configuration | ConvertTo-Json | Out-File -FilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference
}

#endregion Internal Functions