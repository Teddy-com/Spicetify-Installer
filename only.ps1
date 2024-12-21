# Ensure script stops on errors
$ErrorActionPreference = 'Stop'

# Set the security protocol for web requests (not needed in this case but good practice)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region Variables
# Define the path for Spicetify installation
$spicetifyFolderPath = "$env:LOCALAPPDATA\spicetify"
$spicetifyOldFolderPath = "$HOME\spicetify-cli"

# Define the path where you have already downloaded Spicetify (adjust this to your needs)
$spicetifySourceFolder = "C:\path\to\spicetify"  # <-- Change this to the actual folder where you have Spicetify files

#endregion Variables

#region Functions
# Function to output success messages
function Write-Success {
    [CmdletBinding()]
    param ()
    process {
        Write-Host -Object ' > OK' -ForegroundColor 'Green'
    }
}

# Function to output error messages
function Write-Unsuccess {
    [CmdletBinding()]
    param ()
    process {
        Write-Host -Object ' > ERROR' -ForegroundColor 'Red'
    }
}

# Function to check if the script is being run as administrator
function Test-Admin {
    [CmdletBinding()]
    param ()
    begin {
        Write-Host -Object "Checking if the script is not being run as administrator..." -NoNewline
    }
    process {
        $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        -not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
}

# Function to check PowerShell version
function Test-PowerShellVersion {
    [CmdletBinding()]
    param ()
    begin {
        $PSMinVersion = [version]'5.1'
    }
    process {
        Write-Host -Object 'Checking if your PowerShell version is compatible...' -NoNewline
        $PSVersionTable.PSVersion -ge $PSMinVersion
    }
}

# Function to move old Spicetify folder (if exists)
function Move-OldSpicetifyFolder {
    [CmdletBinding()]
    param ()
    process {
        if (Test-Path -Path $spicetifyOldFolderPath) {
            Write-Host -Object 'Moving the old spicetify folder...' -NoNewline
            Copy-Item -Path "$spicetifyOldFolderPath\*" -Destination $spicetifyFolderPath -Recurse -Force
            Remove-Item -Path $spicetifyOldFolderPath -Recurse -Force
            Write-Success
        }
    }
}

# Function to add Spicetify to the PATH environment variable
function Add-SpicetifyToPath {
    [CmdletBinding()]
    param ()
    begin {
        Write-Host -Object 'Making spicetify available in the PATH...' -NoNewline
        $user = [EnvironmentVariableTarget]::User
        $path = [Environment]::GetEnvironmentVariable('PATH', $user)
    }
    process {
        $path = $path -replace "$([regex]::Escape($spicetifyOldFolderPath))\\*;*", ''
        if ($path -notlike "*$spicetifyFolderPath*") {
            $path = "$path;$spicetifyFolderPath"
        }
    }
    end {
        [Environment]::SetEnvironmentVariable('PATH', $path, $user)
        $env:PATH = $path
        Write-Success
    }
}

# Function to install Spicetify from an existing folder
function Install-Spicetify {
    [CmdletBinding()]
    param ()
    begin {
        Write-Host -Object 'Installing spicetify...'
    }
    process {
        # Check if the source folder exists
        if (-not (Test-Path -Path $spicetifySourceFolder)) {
            Write-Unsuccess
            Write-Host -Object "Spicetify source folder does not exist at: $spicetifySourceFolder" -ForegroundColor 'Red'
            exit
        }
        
        # Copy Spicetify files to the target folder
        Write-Host -Object 'Copying spicetify files...' -NoNewline
        Copy-Item -Path "$spicetifySourceFolder\*" -Destination $spicetifyFolderPath -Recurse -Force
        Write-Success

        # Add Spicetify to PATH
        Add-SpicetifyToPath
    }
    end {
        Write-Host -Object 'spicetify was successfully installed!' -ForegroundColor 'Green'
    }
}
#endregion Functions

#region Main
#region Checks
if (-not (Test-PowerShellVersion)) {
    Write-Unsuccess
    Write-Warning -Message 'PowerShell 5.1 or higher is required to run this script'
    Write-Warning -Message "You are running PowerShell $($PSVersionTable.PSVersion)"
    Write-Host -Object 'PowerShell 5.1 install guide:'
    Write-Host -Object 'https://learn.microsoft.com/skypeforbusiness/set-up-your-computer-for-windows-powershell/download-and-install-windows-powershell-5-1'
    Write-Host -Object 'PowerShell 7 install guide:'
    Write-Host -Object 'https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows'
    Pause
    exit
} else {
    Write-Success
}

if (-not (Test-Admin)) {
    Write-Unsuccess
    Write-Warning -Message "The script was run as administrator. This can result in problems with the installation process or unexpected behavior. Do not continue if you do not know what you are doing."
    $Host.UI.RawUI.Flushinputbuffer()
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @(
        (New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Abort installation.'),
        (New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Resume installation.')
    )
    $choice = $Host.UI.PromptForChoice('', 'Do you want to abort the installation process?', $choices, 0)
    if ($choice -eq 0) {
        Write-Host -Object 'spicetify installation aborted' -ForegroundColor 'Yellow'
        Pause
        exit
    }
} else {
    Write-Success
}
#endregion Checks

#region Spicetify Installation
Move-OldSpicetifyFolder
Install-Spicetify
Write-Host -Object "`nRun" -NoNewline
Write-Host -Object ' spicetify -h ' -NoNewline -ForegroundColor 'Cyan'
Write-Host -Object 'to get started'
#endregion Spicetify
#endregion Main
