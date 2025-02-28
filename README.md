# DatedMail

This PowerShell module is intended to provide a simple way for expiring email addresses.
Ultimately this should reduce spam as mail addresses are only available for a certain amount of time.

## Prerequisites

  * PowerShell 6.0 and later
  * A mailbox with [sieve](http://sieve.info/) support
  * Support for the sieve extensions `envelope` and `reject`

## Usage

  1. Start a new PowerShell session

  2. Install the module from the PSGallery using `Install-Module` or `Install-PSResource`
    
```pwsh
PS> Install-PSResource DatedMail
```
  3. Initialize the module by providing a configuration. The default configuration is persisted at `$env::HOME/.config/DatedMail/DatedMailConfig.json` and must provide the email address where valid messages should be forwarded to, as well as the path to the sieve filter file.
    
```pwsh
PS> Initialize-DatedMailConfiguration -SieveFilterPath /home/marvin/users/temp/sieve/datedmail.filter -ForwardingEmailAddress marvin@example.org -MailAddressDomin "example.org" -MailAddressPrefix "temp+"
```
  4. Create a new expiring mail address. This will update the sieve filter and forward all mails received for the new address to the defined forwarding address. The created mail address will be available at the `ExportFilePath` to be used by other scripts. Alternatively it can also be returned to STDOUT.
    
```pwsh
PS> New-DatedMailAddress -ValidDays 7 -ExportFilePath $env:Home/DatedMail
```
  5. Create a timer to periodically call `Update-DatedMailAddress` (e.g. hourly). This will remove expired email addreses from the configuration and update the sieve filter accordingly. Thus mails for expired mail addresses will be rejected and bounced back to the sender with the error message `550 Invalid or expired recipient address`.
  Make sure that the sieve filter specified by the script is the one active on the mailbox.