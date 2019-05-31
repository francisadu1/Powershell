<#
	Script Name   : Microsoft.PowerShell_profile.ps1
	Author        : Luke Leigh
	Created       : 16/03/2019
	Notes         : This script has been created in order pre-configure the following setting:-
					- Shell Title - Rebranded
					- Shell Dimensions configured to 170 Width x 45 Height
					- Buffer configured to 9000 lines
					- Creates GitHub Local Repository PSDrives and Onedrive PSDrives

	EG
	Name                  Root
	----                  ----
	blogsite              C:\GitRepos\blogsite
	CiscoMeraki           C:\GitRepos\CiscoMeraki
	MSPTech               C:\GitRepos\MSPTech
	OneDrive              C:\Users\Luke\OneDrive
	PowerRepo             C:\GitRepos\PowerRepo

- Sets starting file path to Scripts folder on ScriptsDrive
- Loads the following Functions

	CommandType     Name
	-----------     ----
	Function        Get-Appointments
	Function        Get-ContainedCommand
	Function        Get-Password
	Function        Get-PatchTue
	Function        Get-ScriptDirectory
	Function        LoadProfile
	Function        New-GitDrives
	Function        New-Greeting
	Function        New-ObjectToHashTable
	Function        New-PSDrives
	Function        Save-Password
	Function        Select-FolderLocation
	Function        Show-IsAdminOrNot
	Function        Show-ProfileFunctions
	Function        Show-PSDrive
	Function        Stop-Outlook
	Function        Test-IsAdmin

Displays
- whether or not running as Administrator in the WindowTitle
- the Date and Time in the Console Window
- a Greeting based on day of week
- whether or not running as Administrator in the Console Window

When run from Elevated Prompt
- Preconfigures Executionpolicy settings per PowerShell Process Unrestricted
(un-necessary to configure execution policy manually
each new PowerShell session, is configured at run and disposed of on exit)
- Amend PSModulePath variable to include 'OneDrive\PowerShellModules'
- Configure LocalHost TrustedHosts value
- Measures script running performance and displays time upon completion

#>

#--------------------
# Start
$Stopwatch = [system.diagnostics.stopwatch]::startNew()

#--------------------
# Script Functions



# Function        Get-ContainedCommand
function Get-ContainedCommand
{
   param
   (
       [Parameter(Mandatory)][string]
       $Path,

       [string][ValidateSet('FunctionDefinition','Command' )]
       $ItemType
   )

   $Token = $Err = $null
   $ast = [Management.Automation.Language.Parser]::ParseFile( $Path, [ref] $Token, [ref] $Err)

   $ast.FindAll({ $args[0].GetType(). Name -eq "${ItemType}Ast" }, $true )

}

# Function        Get-Password
function Get-Password {
	<#
	.EXAMPLE
	$user = Get-Password -Label UserName
	$pass = Get-Password -Label password

	.OUTPUTS
	$user | Format-List

	.OUTPUTS
	Label           : UserName
	EncryptedString : domain\administrator

	.OUTPUTS
	$pass | Format-List
	Label           : password
	EncryptedString : SomeSecretPassword

	.OUTPUTS
	$user.EncryptedString
	domain\administrator

	.OUTPUTS
	$pass.EncryptedString
	SomeSecretPassword

	#>
	param([Parameter(Mandatory)]
	[string]$Label)
	$directoryPath = Select-FolderLocation
	if (![string]::IsNullOrEmpty($directoryPath)) {
		Write-Host "You selected the directory: $directoryPath"
	}
	$filePath = "$directoryPath\$Label.txt"
	if (-not (Test-Path -Path $filePath)) {
		throw "The password with Label [$($Label)] was not found!"
	}
	$password = Get-Content -Path $filePath | ConvertTo-SecureString
	$decPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
	[pscustomobject]@{
		Label = $Label
		EncryptedString = $decPassword
	}
}

# Function        Get-PatchTue
function Get-PatchTue {
	<#
	.SYNOPSIS
	Get the Patch Tuesday of a month
	.PARAMETER month
	The month to check
	.PARAMETER year
	The year to check
	.EXAMPLE
	Get-PatchTue -month 6 -year 2015
	.EXAMPLE
	Get-PatchTue June 2015
	#>
	param(
		[string]$month = (get-date).month,
		[string]$year = (get-date).year)
		$firstdayofmonth = [datetime] ([string]$month + "/1/" + [string]$year)
		(0..30 | ForEach-Object {
			$firstdayofmonth.adddays($_)
		} |
		Where-Object {
			$_.dayofweek -like "Tue*"
		})[1]
	}

# Function        Get-ScriptDirectory
function Get-ScriptDirectory {
	Split-Path -Parent $PSCommandPath
}

# Function        LoadProfile
function LoadProfile {
	@(
		$Profile.AllUsersAllHosts,
		$Profile.AllUsersCurrentHost,
		$Profile.CurrentUserAllHosts,
		$Profile.CurrentUserCurrentHost
		) |
		ForEach-Object {
			if(Test-Path $_){
				Write-Verbose "Running $_"
				. $_
			}
		}
	}

# Function        New-GitDrives
function New-GitDrives {
	$PSRootFolder = Select-FolderLocation
	$Exist = Test-Path -Path $PSRootFolder
	if ($Exist = $true) {
		$PSDrivePaths = Get-ChildItem -Path "$PSRootFolder\"
		foreach ($item in $PSDrivePaths) {
			$paths = Test-Path -Path $item.FullName
			if ($paths = $true) {
				New-PSDrive -Name $item.Name -PSProvider "FileSystem" -Root $item.FullName
			}
		}
	}
}

# Function        New-Greeting
function New-Greeting {
	$Today = $(Get-Date)
	Write-Host "   Day of Week  -"$Today.DayOfWeek " - Today's Date -"$Today.ToShortDateString() "- Current Time -"$Today.ToShortTimeString()
	Switch ($Today.dayofweek)
	{
		Monday { Write-host "   Don't want to work today" }
		Friday { Write-host "   Almost the weekend" }
		Saturday { Write-host "   Everyone loves a Saturday ;-)" }
		Sunday { Write-host "   A good day to rest, or so I hear." }
		Default { Write-host "   Business as usual." }
	}
}

# Function        New-ObjectToHashTable
function New-ObjectToHashTable{
	param([
		Parameter(Mandatory ,ValueFromPipeline)]
		$object)
		process	{
			$object |
			Get-Member -MemberType *Property |
			Select-Object -ExpandProperty Name |
			Sort-Object |
			ForEach-Object {[PSCustomObject ]@{
				Item = $_
				Value = $object. $_
			}
		}
	}
}

# Function        New-PSDrives
function New-PSDrives {
	$PSRootFolder = Select-FolderLocation
	$PSDrivePaths = Get-ChildItem -Path "$PSRootFolder\"
	foreach ($item in $PSDrivePaths) {
		$paths = Test-Path -Path $item.FullName
		if ($paths = $true) {
			New-PSDrive -Name $item.Name -PSProvider "FileSystem" -Root $item.FullName
		}
	}
}

# Function        Save-Password
function Save-Password {
	<# Example

	.EXAMPLE
	Save-Password -Label UserName

	.EXAMPLE
	Save-Password -Label Password

	#>
	param([Parameter(Mandatory)]
	[string]$Label)
	$securePassword = Read-host -Prompt 'Input password' -AsSecureString | ConvertFrom-SecureString
	$directoryPath = Select-FolderLocation
	if (![string]::IsNullOrEmpty($directoryPath)) {
		Write-Host "You selected the directory: $directoryPath"
	}
	else {
		"You did not select a directory."
	}
	$securePassword | Out-File -FilePath "$directoryPath\$Label.txt"
}

# Function        Select-FolderLocation
function Select-FolderLocation {
    <#
        Example.
        $directoryPath = Select-FolderLocation
        if (![string]::IsNullOrEmpty($directoryPath)) {
            Write-Host "You selected the directory: $directoryPath"
        }
        else {
            "You did not select a directory."
        }
    #>
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = "C:\"
    $browse.ShowNewFolderButton = $true
    $browse.Description = "Select a directory for your report"
    $loop = $true
    while ($loop) {
        if ($browse.ShowDialog() -eq "OK") {
            $loop = $false
        }
        else {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if ($res -eq "Cancel") {
                #Ends script
                return
            }
        }
    }
    $browse.SelectedPath
    $browse.Dispose()
}

# Function        Show-IsAdminOrNot
function Show-IsAdminOrNot {
	$IsAdmin = Test-IsAdmin
	if ( $IsAdmin -eq "False") {
		Write-Warning -Message "Admin Privileges!"
	}
	else {
		Write-Warning -Message "User Privileges"
	}
}

# Function        Show-ProfileFunctions
function Show-ProfileFunctions {
	$Path = $profile
	$functionNames = Get-ContainedCommand $Path -ItemType FunctionDefinition |	Select-Object -ExpandProperty Name
	$functionNames | Sort-Object
}

# Function        Show-PSDrive
	function Show-PSDrive {
		Get-PSDrive | Format-Table -AutoSize
	}

# Function        Stop-Outlook
function Stop-Outlook {
	$OutlookRunning = Get-Process -ProcessName "Outlook"
	if ($OutlookRunning = $true) {
		Stop-Process -ProcessName Outlook
	}
}

# Function        Test-IsAdmin
function Test-IsAdmin {
	<#
	.Synopsis
	Tests if the user is an administrator
	.Description
	Returns true if a user is an administrator, false if the user is not an administrator
	.Example
	Test-IsAdmin
	#>
	$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}



#--------------------
# Display running as Administrator in WindowTitle
if(Test-IsAdmin) {
	Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Force
	$PatchTue = Get-PatchTue -month (Get-Date).Month -year (Get-Date).Year
	if ((get-date).ToShortDateString() -eq ($PatchTue).ToShortDateString()) {
		Update-Help -Force
	}
	$TrustedHosts = Get-Item WSMAN:\localhost\Client\TrustedHosts | Select-Object -Property *
	if ($TrustedHosts.Value = $false) {
		Set-Item WSMAN:\localhost\Client\TrustedHosts -value *
	}
	$host.UI.RawUI.WindowTitle = "$($env:USERNAME) Elevated Shell"
}
else{
	$host.UI.RawUI.WindowTitle = "$($env:USERNAME) Non-elevated Shell"
}

#--------------------
# Configure Powershell Console Window Size/Preferences
$console = $host.UI.RawUI
$buffer = $console.BufferSize
$buffer.Width = 170
$buffer.Height = 9000
$console.BufferSize = $buffer
$size = $console.WindowSize
$size.Width = 170
$size.Height = 45
$console.WindowSize = $size



& "$PSScriptRoot\Connect-Office365Services.ps1"

#--------------------
# Fresh Start
# Clear-Host


# Staff Variables
$FilePath = 'C:\GitRepos\PowerShellScripts\CapitalSupport\Documents\Staff\Staff-Teams.csv'
$Staff = Get-Content -LiteralPath $FilePath | ConvertFrom-Csv
$1stLine = $Staff | Select-Object -Property * | Where-Object { $_.Team -like '*1st*' }
$2ndLine = $Staff | Select-Object -Property * | Where-Object { $_.Team -like '*2nd*' }
$3rdLine = $Staff | Select-Object -Property * | Where-Object { $_.Team -like '*3rd*' }
$DaMan = $Staff | Select-Object -Property * | Where-Object { $_.Team -like '*Da*' }


#--------------------
# Profile Starts here!
Show-IsAdminOrNot
Write-Host ""
New-Greeting
	# Write-Host ""
	# Write-Host "The following Functions are now available in this session"
	# Write-Host ""
	# Show-ProfileFunctions
Write-Host ""
Set-Location -Path C:\GitRepos


#--------------------
# Display Profile Load time and Stop the timer
# Write-Host "Personal Profile took" $Stopwatch.Elapsed.Milliseconds"ms."
$Stopwatch.Stop()
# End --------------#>
