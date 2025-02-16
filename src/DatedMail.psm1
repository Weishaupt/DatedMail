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
        [string] $MailAddressPrefix = "temp+"
    )

    $config = @{
        Addresses = @();
        MailPrefix = $MailAddressPrefix;
        ForwardingEmailAddress = $ForwardingEmailAddress;
        SieveFilterPath = $SieveFilterPath;
    }

    $isValidConfiguration = Test-Configuration -Configuration $config

    if($isValidConfiguration -eq $true) {
        Export-Configuration -Configuration $config -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference
    } else {
        Write-Warning "No configuration was written. Please check the error message and try again."
    }
}

<#
.SYNOPSIS
Adds a new expiring e-mail address to the database

.PARAMETER DaysToExpire
Specifies how many days the new generated address should be valid for.

.PARAMETER DatedMailExportFilePath
Specifies the path to which to export the new email address. The export
will be a text file with only one line consisting of the generated mail address

.PARAMETER ConfigurationFilePath
The configuration path for the configuration. By default a configuration is written
to the .config folder in the users home directory at ~/.config/DatedMail/DatedMailConfig.json

.PARAMETER ReturnMailAddress
Specifies whether the mail address should be returned on STDOUT instead of being written to file.

.EXAMPLE
PS> New-DatedMailAddress -DaysToExpire 7 -DatedMailExportFilePath $env:Home/DatedMail

Creates a new mail address valid for exactly 7 days and exports the address to the file
DatedMail in the users home directory.

#>
function New-DatedMailAddress {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([void], ParameterSetName="File")]
    [OutputType([string], ParameterSetName="Console")]
    param (
        [Parameter(Mandatory=$false)]
        [Int64] $DaysToExpire = 10,
        [Parameter(Mandatory=$false, ParameterSetName="File")]
        [string] $DatedMailExportFilePath = "",
        [Parameter(Mandatory=$false)]
        [string] $ConfigurationFilePath = $DefaultConfigurationFilePath,
        [Parameter(Mandatory=$false, ParameterSetName="Console")]
        [System.Management.Automation.SwitchParameter]$ReturnMailAddress
    )

    if(![String]::IsNullOrWhiteSpace($DatedMailExportFilePath) -and (Test-Path -Path $DatedMailExportFilePath -IsValid -PathType Leaf) -eq $false) {
        Write-Error "The path $DatedMailExportFilePath is invalid. Please provide a valid path and try again."
        throw("Invalid Path")
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

    $DatedMail = [PSCustomObject]@{
        Address = $datedAddress;
        ExpiresOn = [datetime]::Now.AddDays($DaysToExpire).ToString("s")
    }

    Write-Verbose "Generated address '$($DatedMail.Address)' will expire on $($DatedMail.ExpiresOn)"

    $config.Addresses += $DatedMail

    Write-Debug "Adding generated address to existing configuration"
    Export-Configuration -Configuration $config -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference

    Write-Debug "Updating the existing configuration"
    Update-DatedMailAddress -ConfigurationFilePath $ConfigurationFilePath -WhatIf:$WhatIfPreference -ForceSieveFilterUpdate

    if(![String]::IsNullOrWhiteSpace($DatedMailExportFilePath)) {
        Write-Debug "Exporting email address to $DatedMailExportFilePath"
        try {
            $DatedMail.Address | Out-File -FilePath $DatedMailExportFilePath -WhatIf:$WhatIfPreference -NoNewline -ErrorAction Stop
        }
        catch {
            $ex = New-Object -TypeName System.ApplicationException -ArgumentList "Unable to export address to the given file path $DatedMailExportFilePath"
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
        $expiry = [datetime]::Parse($entry.ExpiresOn)
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
        $config = Get-Content -Path $ConfigurationFilePath | ConvertFrom-Json
    } else {
        $ex = New-Object -TypeName System.ApplicationException -ArgumentList "No configuration file found at the given location $($ConfigurationFilePath). Please run Initialize-Configuration or specify the correct location for the configuration file."
        throw($ex)
    }

    $isValidConfiguration = Test-Configuration -Configuration $config

    if($isValidConfiguration -eq $false) {
        $AppException = New-Object System.ApplicationException -ArgumentList "Invalid configuration detected. Check the configuration or run Initialize-Configuration again."
        throw($AppException)
    }

    return $config
}

function Test-Configuration {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)]
        [object[]] $Configuration
    )
    # Prefix
    if($Configuration.MailPrefix -notlike "*+") {
        Write-Error "Mail Address Prefix must end with '+' to allow for Plus-Adressing"
        return $false
    }

    # Sieve Filter Path
    if ((Test-Path $Configuration.SieveFilterPath -PathType Leaf -IsValid) -eq $false) {
        Write-Error "Path to Sieve filter file could not be validated. Please check and try again."
        return $false
    }

    # Forwarding Mail Address
    try {
        New-Object -TypeName "MailAddress" -ArgumentList @($Configuration.ForwardingEmailAddress) | Out-Null
    } catch [FormatException] {
        Write-Error "Invalid forwarding address given: $($Configuration.ForwardingEmailAddress)"
        return $false
    }

    return $true

}

function Export-Configuration {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([void])]
    param(
        [Parameter(Mandatory=$true)]
        [object[]] $Configuration,
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