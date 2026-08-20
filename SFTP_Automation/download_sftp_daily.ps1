# ============================================================
# Daily SFTP downloader to OneDrive
#
# Server: sftp.nx-schweiz.ch
# User: spi-has
#
# What it does:
# - Connects to the SFTP server
# - Checks files in the remote root folder "/"
# - Downloads only files that are not already downloaded
# - Checks both:
#     1. downloaded-files.json tracking file
#     2. existing files already inside the archive folder
# - Saves every downloaded file into one flat archive folder
# - Creates logs
#
# Requires:
# - WinSCP installed
# ============================================================

$ErrorActionPreference = "Stop"

# ----------------------------
# SERVER CONFIG
# ----------------------------

$HostName = "sftp.nx-schweiz.ch"
$PortNumber = 22
$UserName = "spi-has"

# Put your password directly here
$Password = ".LFior4aFwhs@rNRvrE-fBsC"

# Server files are in the root folder
$RemoteRoot = "/"

# ----------------------------
# LOCAL ONEDRIVE CONFIG
# ----------------------------

# Downloads go into one folder next to this script:
# SFTP_Automation\Downloaded_Server_Files_All
$DownloadRoot = Join-Path $PSScriptRoot "Downloaded_Server_Files_All"

# Old year/month folder location. If it still exists, the script copies any
# missing files from there into the flat archive before downloading.
$LegacyMonthRoot = Join-Path $PSScriptRoot "Downloaded_Server_Files"

# Logs go here:
# SFTP_Automation\_automation_logs
$LogFolder = Join-Path $PSScriptRoot "_automation_logs"

# Leave empty to download every file type.
# Example if you only want CSV:
# $AllowedExtensions = @(".csv")
$AllowedExtensions = @()

# You said the files are in the server root, so keep this false.
$IncludeSubfolders = $false

# Keep false for normal use.
# Set true once only if you want to mark current server files as already known without downloading them.
$BaselineOnly = $false

# ----------------------------
# FIND WINSCP
# ----------------------------

$WinScpDllPaths = @(
    "C:\Program Files (x86)\WinSCP\WinSCPnet.dll",
    "C:\Program Files\WinSCP\WinSCPnet.dll"
)

$WinScpDll = $WinScpDllPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (!$WinScpDll) {
    throw "WinSCPnet.dll was not found. Install WinSCP first, then run this script again."
}

Add-Type -Path $WinScpDll

# ----------------------------
# PREPARE FOLDERS AND FILES
# ----------------------------

New-Item -ItemType Directory -Force -Path $DownloadRoot | Out-Null
New-Item -ItemType Directory -Force -Path $LogFolder | Out-Null

$RunStamp = Get-Date -Format "yyyyMMdd-HHmmss"

$TranscriptLog = Join-Path $LogFolder "download-run-$RunStamp.log"
$SessionLog = Join-Path $LogFolder "winscp-session-$RunStamp.log"
$StateFile = Join-Path $LogFolder "downloaded-files.json"
$LockFile = Join-Path $LogFolder "download.lock"

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------

function Save-DownloadedState {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$DownloadedTable,

        [Parameter(Mandatory = $true)]
        [string]$StateFilePath
    )

    $StateArray = @()

    foreach ($Key in $DownloadedTable.Keys) {
        $StateArray += [string]$Key
    }

    $StateArray |
        Sort-Object |
        ConvertTo-Json |
        Set-Content -Path $StateFilePath -Encoding UTF8
}

function Add-ToHashTable {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Table,

        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    if (!$Table.ContainsKey($Key)) {
        $Table[$Key] = $true
    }
}

function Test-InHashTable {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Table,

        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    return $Table.ContainsKey($Key)
}

# ----------------------------
# START LOG
# ----------------------------

Start-Transcript -Path $TranscriptLog -Append | Out-Null

try {
    Write-Host "============================================================"
    Write-Host "SFTP daily download started: $(Get-Date)"
    Write-Host "Server: $HostName"
    Write-Host "Remote folder: $RemoteRoot"
    Write-Host "Archive folder: $DownloadRoot"
    Write-Host "Baseline only mode: $BaselineOnly"
    Write-Host "============================================================"

    if ([string]::IsNullOrWhiteSpace($Password) -or $Password -eq "PASTE_YOUR_PASSWORD_HERE") {
        throw "Password is missing. Edit the script and put the SFTP password in the `$Password variable."
    }

    # Prevent two runs at the same time
    if (Test-Path $LockFile) {
        throw "Another download job seems to be running. Lock file exists: $LockFile"
    }

    New-Item -ItemType File -Path $LockFile -Force | Out-Null

    # ----------------------------
    # LOAD TRACKING FILE
    # ----------------------------

    $Downloaded = @{}

    if (Test-Path $StateFile) {
        try {
            $ExistingStateRaw = Get-Content $StateFile -Raw

            if (![string]::IsNullOrWhiteSpace($ExistingStateRaw)) {
                $ExistingState = $ExistingStateRaw | ConvertFrom-Json

                foreach ($Item in @($ExistingState)) {
                    if (![string]::IsNullOrWhiteSpace($Item)) {
                        Add-ToHashTable -Table $Downloaded -Key ([string]$Item)
                    }
                }
            }
        }
        catch {
            Write-Host "WARNING: Could not read downloaded-files.json. The script will rely on local file scanning."
            Write-Host $_.Exception.Message
        }
    }

    # ----------------------------
    # SCAN LOCAL ARCHIVE FOLDER
    # ----------------------------

    $ExistingLocalFileNames = @{}
    $LegacyBackfillCount = 0

    Get-ChildItem -Path $DownloadRoot -File -ErrorAction SilentlyContinue | ForEach-Object {
        $LocalNameKey = $_.Name.ToLowerInvariant()

        if (!$ExistingLocalFileNames.ContainsKey($LocalNameKey)) {
            $ExistingLocalFileNames[$LocalNameKey] = $_.FullName
        }
    }

    if (Test-Path $LegacyMonthRoot) {
        Get-ChildItem -Path $LegacyMonthRoot -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $LocalNameKey = $_.Name.ToLowerInvariant()

            if (!$ExistingLocalFileNames.ContainsKey($LocalNameKey)) {
                $DestinationPath = Join-Path $DownloadRoot $_.Name
                Copy-Item -Path $_.FullName -Destination $DestinationPath
                $ExistingLocalFileNames[$LocalNameKey] = $DestinationPath
                $LegacyBackfillCount++
            }
        }
    }

    Write-Host "Existing archive files found: $($ExistingLocalFileNames.Count)"
    Write-Host "Files copied from old month folder into archive: $LegacyBackfillCount"
    Write-Host "Already tracked server files: $($Downloaded.Count)"

    # ----------------------------
    # CONNECT TO SFTP
    # ----------------------------

    $SessionOptions = New-Object WinSCP.SessionOptions
    $SessionOptions.Protocol = [WinSCP.Protocol]::Sftp
    $SessionOptions.HostName = $HostName
    $SessionOptions.PortNumber = $PortNumber
    $SessionOptions.UserName = $UserName
    $SessionOptions.Password = $Password

    # Your earlier run connected successfully with this, so keep it.
    try {
        $SessionOptions.SshHostKeyPolicy = [WinSCP.SshHostKeyPolicy]::AcceptNew
    }
    catch {
        throw "Your WinSCP version does not support SshHostKeyPolicy AcceptNew. Update WinSCP to the latest version."
    }

    $Session = New-Object WinSCP.Session
    $Session.SessionLogPath = $SessionLog

    $TransferOptions = New-Object WinSCP.TransferOptions
    $TransferOptions.TransferMode = [WinSCP.TransferMode]::Binary

    try {
        $Session.Open($SessionOptions)
        Write-Host "Connected successfully."

        # ----------------------------
        # GET SERVER FILES
        # ----------------------------

        if ($IncludeSubfolders) {
            $RemoteFiles = $Session.EnumerateRemoteFiles(
                $RemoteRoot,
                "*",
                [WinSCP.EnumerationOptions]::AllDirectories
            )
        }
        else {
            $RemoteDirectory = $Session.ListDirectory($RemoteRoot)
            $RemoteFiles = $RemoteDirectory.Files
        }

        $DownloadCount = 0
        $SkipTrackedCount = 0
        $SkipLocalExistsCount = 0
        $IgnoredCount = 0
        $BaselineCount = 0
        $ErrorCount = 0

        foreach ($RemoteFile in $RemoteFiles) {
            try {
                if ($RemoteFile.IsDirectory) {
                    continue
                }

                if ($RemoteFile.Name -eq "." -or $RemoteFile.Name -eq "..") {
                    continue
                }

                $Extension = [System.IO.Path]::GetExtension($RemoteFile.Name).ToLowerInvariant()

                if ($AllowedExtensions.Count -gt 0 -and $AllowedExtensions -notcontains $Extension) {
                    $IgnoredCount++
                    continue
                }

                # This key identifies the exact server file version.
                # If same path + same size + same modified date exists, skip it.
                $FileKey = "$($RemoteFile.FullName)|$($RemoteFile.Length)|$($RemoteFile.LastWriteTime.ToUniversalTime().ToString('o'))"

                $RemoteFileNameKey = $RemoteFile.Name.ToLowerInvariant()

                # 1. Skip if already tracked and the local file still exists.
                # If a local file was deleted or moved outside this download folder, recover it.
                if (Test-InHashTable -Table $Downloaded -Key $FileKey) {
                    if ($ExistingLocalFileNames.ContainsKey($RemoteFileNameKey)) {
                        $SkipTrackedCount++
                        continue
                    }

                    Write-Host "Tracked but local file is missing, downloading again: $($RemoteFile.Name)"
                }

                # 2. Skip if same filename already exists locally anywhere in Downloaded_Server_Files
                if ($ExistingLocalFileNames.ContainsKey($RemoteFileNameKey)) {
                    Write-Host "Skipping because local file already exists: $($RemoteFile.Name)"

                    Add-ToHashTable -Table $Downloaded -Key $FileKey
                    Save-DownloadedState -DownloadedTable $Downloaded -StateFilePath $StateFile

                    $SkipLocalExistsCount++
                    continue
                }

                # 3. Baseline mode: mark as known but do not download
                if ($BaselineOnly) {
                    Write-Host "Baseline only - marking as known without downloading: $($RemoteFile.Name)"

                    Add-ToHashTable -Table $Downloaded -Key $FileKey
                    Save-DownloadedState -DownloadedTable $Downloaded -StateFilePath $StateFile

                    $BaselineCount++
                    continue
                }

                $LocalFilePath = Join-Path $DownloadRoot $RemoteFile.Name
                $TempFilePath = "$LocalFilePath.partial"

                # Final overwrite protection
                if (Test-Path $LocalFilePath) {
                    Write-Host "Skipping because destination file already exists: $LocalFilePath"

                    Add-ToHashTable -Table $Downloaded -Key $FileKey
                    Save-DownloadedState -DownloadedTable $Downloaded -StateFilePath $StateFile

                    $SkipLocalExistsCount++
                    continue
                }

                # Remove old partial file if present
                if (Test-Path $TempFilePath) {
                    Remove-Item $TempFilePath -Force
                }

                Write-Host "Downloading: $($RemoteFile.FullName)"
                Write-Host "To: $LocalFilePath"

                $Result = $Session.GetFiles(
                    $RemoteFile.FullName,
                    $TempFilePath,
                    $false,
                    $TransferOptions
                )

                $Result.Check()

                Move-Item -Path $TempFilePath -Destination $LocalFilePath -Force

                Add-ToHashTable -Table $Downloaded -Key $FileKey
                if (!$ExistingLocalFileNames.ContainsKey($RemoteFileNameKey)) {
                    $ExistingLocalFileNames[$RemoteFileNameKey] = $LocalFilePath
                }

                Save-DownloadedState -DownloadedTable $Downloaded -StateFilePath $StateFile

                $DownloadCount++
            }
            catch {
                $ErrorCount++
                Write-Host "ERROR while processing file: $($RemoteFile.FullName)"
                Write-Host $_.Exception.Message

                # Continue with next file instead of stopping whole job
                continue
            }
        }

        Write-Host "============================================================"
        Write-Host "Download finished."
        Write-Host "Downloaded new files: $DownloadCount"
        Write-Host "Skipped because already tracked: $SkipTrackedCount"
        Write-Host "Skipped because already in archive folder: $SkipLocalExistsCount"
        Write-Host "Ignored by extension filter: $IgnoredCount"
        Write-Host "Baseline marked without downloading: $BaselineCount"
        Write-Host "Files with errors: $ErrorCount"
        Write-Host "============================================================"
    }
    finally {
        if ($Session) {
            $Session.Dispose()
        }
    }
}
catch {
    Write-Host "MAIN ERROR:"
    Write-Host $_.Exception.Message
    throw
}
finally {
    if (Test-Path $LockFile) {
        Remove-Item $LockFile -Force
    }

    Write-Host "Finished at $(Get-Date)"
    Stop-Transcript | Out-Null
}
