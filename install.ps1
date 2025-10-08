# Requires: PowerShell 5+ on Windows
# Behavior: Elevates, compares versions between installed and packaged bundle, prompts if different, replaces, reports errors

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Admin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName  = 'powershell.exe'
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
        if (-not $scriptPath) { $scriptPath = $MyInvocation.InvocationName }
        $args = @(
            '-NoProfile','-ExecutionPolicy','Bypass','-File',
            ('"{0}"' -f $scriptPath)
        )
        if ($PSBoundParameters.Count -gt 0) {
            $PSBoundParameters.GetEnumerator() | ForEach-Object {
                $args += @($_.Key, $_.Value)
            }
        }
        $psi.Arguments = $args -join ' '
        $psi.Verb = 'RunAs'
        try {
            [Diagnostics.Process]::Start($psi) | Out-Null
        } catch {
            Write-Error "Elevation was denied. Aborting."
        }
        exit
    }
}

function Get-PackageVersionFromXml {
    param(
        [Parameter(Mandatory)]
        [string] $XmlPath
    )
    if (-not (Test-Path -LiteralPath $XmlPath)) {
        return $null
    }
    try {
        [xml]$xml = Get-Content -LiteralPath $XmlPath -Encoding UTF8
        # Prefer ApplicationPackage@AppVersion; fallback to first ComponentEntry@Version
        $appVersion = $xml.ApplicationPackage.AppVersion
        if ($appVersion -and $appVersion.Trim().Length -gt 0) {
            return $appVersion.Trim()
        }
        $componentEntry = $xml.SelectSingleNode('//ComponentEntry[@Version]')
        if ($componentEntry -and $componentEntry.Version) {
            return ($componentEntry.Version).Trim()
        }
    } catch {
        Write-Verbose ("Failed to parse {0}: {1}" -f $XmlPath, $_)
    }
    return $null
}

function Prompt-YesNo {
    param(
        [Parameter(Mandatory)][string] $Message,
        [string] $Title = 'TA-Tools Installer'
    )
    $choices = @(
        New-Object System.Management.Automation.Host.ChoiceDescription '&Yes','Proceed'
        New-Object System.Management.Automation.Host.ChoiceDescription '&No','Cancel'
    )
    $selection = $Host.UI.PromptForChoice($Title, $Message, $choices, 1)
    return ($selection -eq 0)
}

Ensure-Admin

$targetRoot  = Join-Path $env:ProgramFiles 'Autodesk\ApplicationPlugins'
$bundleName  = 'TaTools.bundle'
$targetPath  = Join-Path $targetRoot $bundleName

# Support IExpress payload: if TaTools.bundle.zip is present next to the script,
# expand it to a temporary directory and use it as the source root.
$bundleZip   = Join-Path $PSScriptRoot 'TaTools.bundle.zip'
$tempExtract = $null
if (Test-Path -LiteralPath $bundleZip) {
    $tempExtract = Join-Path $env:TEMP ("TaTools_bundle_extract_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempExtract | Out-Null
    Expand-Archive -Path $bundleZip -DestinationPath $tempExtract -Force
    $sourceRoot  = $tempExtract
} else {
    # Fallback: the bundle directory is expected to be next to the script (SFX/7z flow)
    $sourceRoot  = $PSScriptRoot
}
$sourcePath  = Join-Path $sourceRoot $bundleName

# Logging helper and transcript
$eventLogPath   = Join-Path $env:TEMP 'TA-Tools-Installer.events.log'
$transcriptPath = Join-Path $env:TEMP 'TA-Tools-Installer.transcript.log'
function Write-Log {
    param([Parameter(Mandatory)][string]$Message)
    try {
        $line = "[{0}] {1}" -f (Get-Date -Format s), $Message
        $fileStream = [System.IO.File]::Open($eventLogPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        try {
            $writer = New-Object System.IO.StreamWriter($fileStream, [System.Text.Encoding]::UTF8)
            try { $writer.WriteLine($line) } finally { $writer.Dispose() }
        } finally { $fileStream.Dispose() }
    } catch {}
    Write-Host $Message
}
try { Start-Transcript -Path $transcriptPath -Append -ErrorAction SilentlyContinue | Out-Null } catch {}
Write-Log "Installer started."
Write-Log ("Detected source: '{0}'" -f $sourcePath)
Write-Log ("Target path: '{0}'" -f $targetPath)

# Read versions
$sourceXml   = Join-Path $sourcePath 'PackageContents.xml'
$targetXml   = Join-Path $targetPath 'PackageContents.xml'

$sourceVersion = Get-PackageVersionFromXml -XmlPath $sourceXml
$targetVersion = Get-PackageVersionFromXml -XmlPath $targetXml

Write-Log "Packaged version: $sourceVersion"
if (Test-Path -LiteralPath $targetPath) {
    Write-Log "Installed path: $targetPath"
    Write-Log "Installed version: $targetVersion"
} else {
    Write-Log "Installed path: $targetPath (not found)"
}

$shouldInstall = $true

if (Test-Path -LiteralPath $targetPath) {
    if ($sourceVersion -and $targetVersion -and ($sourceVersion -ne $targetVersion)) {
        $shouldInstall = Prompt-YesNo -Message "Installed version ($targetVersion) differs from packaged version ($sourceVersion). Replace it?"
    } elseif ($sourceVersion -and $targetVersion -and ($sourceVersion -eq $targetVersion)) {
        # Same version; do nothing unless user explicitly wants to reinstall
        $shouldInstall = Prompt-YesNo -Message "Installed version equals packaged version ($sourceVersion). Reinstall anyway?"
    } else
        {
        # Could not read one of versions; ask user
        $shouldInstall = Prompt-YesNo -Message "Could not reliably read version. Replace installed bundle?"
    }

    if (-not $shouldInstall) {
        Write-Log "No changes made."
        exit 0
    }
}

try {
    if (Test-Path -LiteralPath $targetPath) {
        Write-Log "Removing existing '$targetPath'..."
        # Try to unlock common file locks by removing read-only attributes
        Get-ChildItem -LiteralPath $targetPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try { $_.Attributes = 'Normal' } catch {}
        }
        Remove-Item -LiteralPath $targetPath -Recurse -Force -ErrorAction Stop
    }

    Write-Log "Copying '$sourcePath' -> '$targetRoot'..."
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Packaged bundle not found at '$sourcePath'."
    }
    if (-not (Test-Path -LiteralPath $targetRoot)) {
        New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
    }
    Copy-Item -LiteralPath $sourcePath -Destination $targetRoot -Recurse -Force -ErrorAction Stop
    Write-Log "Installation completed successfully."
    if ($tempExtract -and (Test-Path -LiteralPath $tempExtract)) {
        Remove-Item -LiteralPath $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
    }
    try { Stop-Transcript | Out-Null } catch {}
    exit 0
}
catch {
    Write-Error ("Installation failed: {0}" -f $_.Exception.Message)
    Write-Log ("Installation failed: {0}" -f $_.Exception.Message)
    if ($tempExtract -and (Test-Path -LiteralPath $tempExtract)) {
        Remove-Item -LiteralPath $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
    }
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}