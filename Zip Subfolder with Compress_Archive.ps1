$scriptVersion = "1802alpha19"
<#
.SYNOPSIS
    Zip SharePoint subfolders from a target library folder, upload ZIP back, log history, and delete the original folders.

.DESCRIPTION
    - Downloads each subfolder (excluding ZZZ_COMPRESSION_HISTORY) under a target folder from SPO to local temp.
    - If a ZIP with matching 6-digit reference exisDts at root level, it downloads and updates it; otherwise creates new ZIP.
    - Uploads ZIP back to the same location.
    - Writes a per-root folder compression log to ZZZ_COMPRESSION_HISTORY.
    - Deletes the original (now compressed) folder.
    - Handles long local paths by truncating filenames with a stable short SHA1 suffix.

.NOTES
    Author: HG33YH - Fabrice Voarick
    Optimised: 2026-01-15
    PnP.PowerShell: v1.12.0 (PowerShell 5.1 compatible)

.CHANGELOG

    1802alpha19 / 01.07.2026
      - CHANGE alpha19: $TempRoot is now created automatically if it does not exist instead of aborting the run.
      - CHANGE alpha19: option D now proposes an optional merge after duplicate ZIP detection.
      - CHANGE alpha19: option D can merge all duplicate ZIP groups in one selected root folder, or all duplicate ZIP groups across every scanned root folder.
      - CHANGE alpha19: duplicate ZIP merge supports more than two ZIPs per reference.
      - CHANGE alpha19: before uploading a ZIP, the script re-checks destination ZIPs and skips/logs the action if an unexpected ZIP with the same 6-digit reference or destination name already exists.

    1802alpha18 / 17.06.2026
      - CHANGE alpha17: added menu option C to check that uncompressed folders and already-compressed ZIP files are stored in the correct 6-digit range folder.
      - CHANGE alpha17: option C is display-only: it does not download, upload, delete or write the detected placement issues to any log file.
      - CHANGE alpha17: option Q (quit) is now default is user is pressing ENTER on menu selection.

    1802alpha16 / 17.06.2026
      - CHANGE alpha16: every folder skipped because it has no standalone 6-digit reference is now logged individually on screen, in the main run log, and in the per-root compression history log.
      - CHANGE alpha16: the per-root compression log is created before the no-reference filtering so skipped folders can be audited even when nothing is downloaded.

    1802alpha15 / 17.06.2026
      - CHANGE alpha15: folders without a standalone 6-digit reference are excluded before processing/download.
      - CHANGE alpha15: added a mandatory reference guard before empty-check, local workspace creation, SharePoint download, ZIP merge, upload and delete.

    1802alpha14 / 03.06.2026
      - FIX alpha14: normalize merge roots so all merged content lands under one consistent top-level folder inside the final ZIP.
      - FIX alpha14: when an extracted ZIP contains a single top-level folder, merge uses that folder's contents instead of creating parallel sibling folders.
      - FIX alpha13: in duplicate ZIP merge, the final merged ZIP is no longer deleted before Add-PnPFile upload.
      - CHANGE alpha13: temporary downloaded duplicate ZIPs now sit in a separate local area from the final ZIP output.
      - CHANGE alpha12: verified compact-path alignment for duplicate and single-ZIP merge workflows.
      - CHANGE alpha12: duplicate ZIP merge prompt is clearer and explicitly offers source folder ZIP name or any existing ZIP name.

    1802alpha / 11.05.2026
      - CHANGE: added 'D' option to detect duplicate references in folders.
      - CHANGE: Make logs ZIP-centric (reference/ZIP as main entity) to avoid confusing 'upload then delete folder' messages.
      - CHANGE: Upload/Delete log lines now explicitly state source folder vs merged ZIP and include [REF <6digits>] prefix.
      - CHANGE: Add conflict-handling strategies during merge (LatestWins, Rename, Fail) using in-memory file inventories.
      - CHANGE: Use staging merge for existing ZIP updates (extract -> merge -> rebuild) so contents merge instead of adding folder container.
      - CHANGE: Apply conflict strategy also when multiple ZIPs exist for same reference by extracting each ZIP separately and merging via in-memory plan.
      - CHANGE: Conflict mode is now a script variable (no command-line argument needed).
                - LatestWins : keep most recent LastWriteTimeUtc; if equal timestamp, keep existing
                - Rename     : keep both by renaming incoming file with _DUPn suffix
                - Fail       : abort merge on first conflict set

    1801beta / 15.01.2026:
      - Fixed server-relative URL building for checkout/checkin.
      - Add-PnPFile now correct site-relative -Folder.
      - Safer temp handling (per run), -WhatIf support, try -Interactive first.
      - Replaced Find-PnPFile for local folder-level zip discovery.
      - Better error handling and disposal in ZIP operations.

    1800 / 05.12.2025:
      - Get-SafeFileName and Get-HashSuffix for <=255 local path constraint.
      1800 / 05.12.2025: new feature,  Read-SpoFolder function will truncate files with path too long automatically.
                    if a file to be downloaded from Sharepoint is too long for download on local machine, function will rename file to fit <255 cars.
                    file will be truncated with a 6 cars hash at the end to ensure uniqueness and length <255
                    example:
                    Screening for LC_25_17928-715877 UTIL 2-TRASTEEL-T2-MS2-U_CCR_SL_CLEAN Clean .msg , will be renamed to
                   Screening for LC_25_17928-715877 UTIL 2-TRASTEEL-T2-MS2-U_CCR_SL_CLE_6ed2c9.msg 
                   (also taking onto account the full path of the folder until we reach the file)

                    resulting no longer onto local folder download errors before compression and zip resent to Sharepoint.
                   see Function Get-HashSuffix and Get-SafeFileName called by Read-SpoFolder.

    1708 / 04.12.2025: no need to process an empty subitem:
    -> we skip it 
    -> we avoid to download an existing zip to update if any
    -> finally we dont raise an error for this on the logs, only a warning reporting that folder is empty and skipped.

    1707: for each subfolder, the ZZZ_Compression history is now populated with compression activities.

    1706: first working version, log is local, no compression history logs populated yet on sharepoint.
       

    we are using module version 1.12, last version compatible with built in Powershell v 5.1 and PS ISE console on windows 11.
    next version is 2.0, only compatible with Powershell 7.x and MS Visual studio code, quite heavy to deploy, we dont need that stuff.
    see https://www.powershellgallery.com/packages/PnP.PowerShell/1.12.0 
    When installing module on a new machine, use this command from a powershell command prompt to force that version for install:
    Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0
    (require local admin rights)
    then execute Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser  ,so powershell has enough rights to execute.

    
    #Disable update version check of the PNP PowerShell module
    #Requires -Version 5.1
    #Requires -Modules PnP.PowerShell
    #$env:PNPPOWERSHELL_UPDATECHECK = "OFF"
#>


param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $ShpDomain = "https://xxx.sharepoint.com"

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $Site ="/sites/GRCH000062_FS" # site-relative for consistency

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $DocLib = "/Shared Documents",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $DocFolder = "/_Archived Files",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $TempRoot = "C:\ProgramData\NA",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $LogFolderName = "ZZZ_COMPRESSION_HISTORY",

    [switch] $ProcessAll,  # If set, skips the interactive prompt and processes all root folders
    [switch] $Quiet        # If set, less console output (log still complete)
    )

begin {
    # ---- SETTINGS / GLOBALS ----
    $ErrorActionPreference = 'Stop'
    $ProgressPreference    = 'SilentlyContinue'  # speeds up lots of cmdlet calls
    $env:PNPPOWERSHELL_UPDATECHECK = "OFF"
    $scriptVersion = "1802alpha19"   # CHANGE alpha19: option D duplicate ZIP merge + safer temp/upload guards
    $ConflictMode = 'LatestWins'
    $dbg = $false
    $SiteUrl = "$ShpDomain$Site"
    $FolderSiteRelativeURL = "$DocLib$DocFolder"                 # site-relative: "/Shared Documents/_Archived Files"
    $FolderServerRelativeURL = "$Site$FolderSiteRelativeURL"     # server-relative: "/sites/.../Shared Documents/_Archived Files"

    # temp structure: create the temp root if missing, then a per-run folder
    if (-not (Test-Path -LiteralPath $TempRoot)) {
        try {
            New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
        }
        catch {
            throw "Temp root '$TempRoot' does not exist and could not be created. Aborting. $($_.Exception.Message)"
        }
    }
    
    #$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    #$RunTemp = Join-Path $TempRoot ("SPOZip_" + $runStamp)
    #New-Item -ItemType Directory -Path $RunTemp | Out-Null
    
    
    # Generate a random 2-character string (letters and numbers)
    $runStamp = -join ((65..90 + 97..122 + 48..57) | Get-Random -Count 2 | ForEach-Object {[char]$_})

    # Example usage:
    $RunTemp = Join-Path $TempRoot $runStamp
    New-Item -ItemType Directory -Path $RunTemp | Out-Null


    # Logs
    $localLogPath = Join-Path $RunTemp ("log_run_$runStamp.txt")
    "timestamp,severity,message" | Out-File -FilePath $localLogPath -Encoding UTF8 -Force
    $subLogPath = $null  # will be created per root folder
}

process

{
    function Write-log  {
        param(
            [Parameter(Mandatory)] [string] $Message,
            [ValidateSet('DEBUG', 'INFO','WARN','ERROR','FATAL')]
            [string] $Severity = "INFO",
            [switch] $ToConsole,
            [switch] $Sub
        )
        $line = ('{0},{1},{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Severity, $Message)
        Add-Content -Path $localLogPath -Value $line
        if ($Sub -and $subLogPath) { Add-Content -Path $subLogPath -Value ("{0},{1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) }
        if ($ToConsole -and -not $Quiet) {
            switch ($Severity) {
                'ERROR' { Write-Host $Message -ForegroundColor Red }
                'WARN'  { Write-Host $Message -ForegroundColor Yellow }
                'DEBUG' { Write-Host $Message -ForegroundColor Gray }
                default { Write-Host $Message }
            }
        }
    }

    function Get-SixDigitReference {
        param(
            [Parameter(Mandatory)] [string] $Name
        )
        # CHANGE alpha15: one authoritative extraction rule for folder/ZIP references.
        # The lookarounds prevent taking 6 digits from inside a longer number.
        if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
        if ($Name -match '(?<!\d)(?<ref>\d{6})(?!\d)') { return $Matches['ref'] }
        return $null
    }


    function Get-SixDigitRangeFromName {
        param(
            [Parameter(Mandatory)] [string] $Name
        )
        # CHANGE alpha17: parse root range folders such as "980000-982999" or "980000 - 982999".
        if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
        if ($Name -match '^\s*(?<start>\d{6})\s*-\s*(?<end>\d{6})\s*$') {
            $startValue = [int]$Matches['start']
            $endValue   = [int]$Matches['end']
            if ($startValue -gt $endValue) { return $null }
            return [pscustomobject]@{
                Start       = $startValue
                End         = $endValue
                StartText   = $Matches['start']
                EndText     = $Matches['end']
                DisplayText = ("{0}-{1}" -f $Matches['start'], $Matches['end'])
            }
        }
        return $null
    }

    function Invoke-PlacementCheck {
        param(
            [Parameter(Mandatory)] [object[]] $RootFolders,
            [Parameter(Mandatory)] [string] $FolderSiteRelativeURL,
            [Parameter(Mandatory)] [string] $FolderServerRelativeURL,
            [Parameter(Mandatory)] $Connection,
            [Parameter(Mandatory)] [string] $LogFolderName
        )

        # CHANGE alpha17: display-only audit. Do not call Write-log here.
        Write-Host "Checking placement of uncompressed folders and already-compressed ZIP files..." -ForegroundColor Cyan
        Write-Host "No download, upload, delete or log write will be performed by this check." -ForegroundColor Cyan

        $scanFolders = @($RootFolders | Where-Object { $_ -and $_.Name -and $_.Name -ne $LogFolderName })
        $issueRows = New-Object System.Collections.Generic.List[object]
        $skippedRangeFolders = New-Object System.Collections.Generic.List[string]
        $checkedItemCount = 0

        foreach ($rf in $scanFolders)
        {
            $rangeName = $rf.Name
            $range = Get-SixDigitRangeFromName -Name $rangeName
            if ($null -eq $range)
            {
                $skippedRangeFolders.Add($rangeName)
                continue
            }

            $subrooturl = "$FolderSiteRelativeURL/$rangeName"

            try
            {
                $foldersAtRoot = @(Get-PnPFolderItem -FolderSiteRelativeUrl $subrooturl -ItemType Folder -Connection $Connection -ErrorAction Stop |
                                  Where-Object { $_.Name -ne $LogFolderName })
            }
            catch
            {
                Write-Host ("Cannot enumerate folders in {0} - {1}" -f $subrooturl, $_.Exception.Message) -ForegroundColor Yellow
                continue
            }

            try
            {
                $filesAtRoot = @(Get-PnPFolderItem -FolderSiteRelativeUrl $subrooturl -ItemType File -Connection $Connection -ErrorAction Stop)
            }
            catch
            {
                Write-Host ("Cannot enumerate files in {0} - {1}" -f $subrooturl, $_.Exception.Message) -ForegroundColor Yellow
                $filesAtRoot = @()
            }

            $zipItems = @($filesAtRoot | Where-Object { $_.Name -match '(?i)\.zip$' })

            $itemsToCheck = @()
            foreach ($folderItem in $foldersAtRoot)
            {
                $itemsToCheck += [pscustomobject]@{
                    ItemType = 'Folder'
                    Name     = $folderItem.Name
                    RawItem  = $folderItem
                }
            }
            foreach ($zipItem in $zipItems)
            {
                $itemsToCheck += [pscustomobject]@{
                    ItemType = 'ZIP'
                    Name     = $zipItem.Name
                    RawItem  = $zipItem
                }
            }

            foreach ($item in $itemsToCheck)
            {
                $checkedItemCount++
                $ref = Get-SixDigitReference -Name $item.Name
                $reason = $null

                if ([string]::IsNullOrWhiteSpace($ref))
                {
                    $reason = 'No standalone 6-digit reference'
                    $refForReport = ''
                }
                else
                {
                    $refValue = [int]$ref
                    $refForReport = $ref
                    if (($refValue -lt $range.Start) -or ($refValue -gt $range.End))
                    {
                        $reason = 'Reference outside expected range'
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($reason))
                {
                    $sr = $null
                    if ($item.RawItem.PSObject.Properties.Match('ServerRelativeUrl').Count -gt 0) { $sr = $item.RawItem.ServerRelativeUrl }
                    if ([string]::IsNullOrWhiteSpace($sr)) { $sr = "$FolderServerRelativeURL/$rangeName/$($item.Name)" }

                    $issueRows.Add([pscustomobject]@{
                        RangeFolder       = $rangeName
                        ExpectedRange     = $range.DisplayText
                        ItemType          = $item.ItemType
                        Ref               = $refForReport
                        ItemName          = $item.Name
                        Reason            = $reason
                        ServerRelativeUrl = $sr
                    })
                }
            }
        }

        Write-Host ""
        Write-Host ("Checked {0} folder/ZIP item(s) across {1} root folder(s)." -f $checkedItemCount, $scanFolders.Count) -ForegroundColor Cyan

        if ($issueRows.Count -eq 0)
        {
            Write-Host "No misplaced folder or ZIP file found." -ForegroundColor Green
        }
        else
        {
            Write-Host ("Misplaced or unverifiable folder/ZIP item(s): {0}" -f $issueRows.Count) -ForegroundColor Yellow
            $issueRows |
                Sort-Object RangeFolder, ItemType, Ref, ItemName |
                Format-Table RangeFolder, ExpectedRange, ItemType, Ref, ItemName, Reason -AutoSize |
                Out-String -Width 300 |
                Write-Host

            Write-Host "Full paths:" -ForegroundColor Yellow
            foreach ($row in ($issueRows | Sort-Object RangeFolder, ItemType, Ref, ItemName))
            {
                Write-Host (" - [{0}] {1} | Ref='{2}' | {3} | {4}" -f $row.RangeFolder, $row.ItemType, $row.Ref, $row.Reason, $row.ServerRelativeUrl)
            }
        }

        if ($skippedRangeFolders.Count -gt 0)
        {
            Write-Host ""
            Write-Host "Root folders skipped because their name is not a valid 6-digit range:" -ForegroundColor Yellow
            foreach ($skipped in ($skippedRangeFolders | Sort-Object))
            {
                Write-Host (" - {0}" -f $skipped)
            }
        }
    }

    function Get-HashSuffix {
        param(
            [Parameter(Mandatory)] [string] $Text,
            [int] $Length = 6
        )
        $sha1 = [System.Security.Cryptography.SHA1]::Create()
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
            $hashBytes = $sha1.ComputeHash($bytes)
            $hex = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
            if ($Length -gt $hex.Length) { $Length = $hex.Length }
            return $hex.Substring(0, $Length)
        } finally {
            $sha1.Dispose()
        }
    }

    function Get-SafeFileName {
        param(
            [Parameter(Mandatory)] [string] $FolderPath,
            [Parameter(Mandatory)] [string] $OriginalName,
            [int] $MaxPathLength = 255
        )

        # $sep = [System.IO.Path]::DirectorySeparatorChar  VS code says variable not used anymore


        $fullOriginal = Join-Path -Path $FolderPath -ChildPath $OriginalName

        if ($fullOriginal.Length -le $MaxPathLength) { return $OriginalName }

        $ext  = [System.IO.Path]::GetExtension($OriginalName)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($OriginalName)
        $roomForName = $MaxPathLength - ($FolderPath.Length + 1)  # +1 for separator

        if ($roomForName -le 8) {
            $suffix = Get-HashSuffix -Text $OriginalName -Length 10
            return "f_$suffix$ext"
        }

        $suffix = Get-HashSuffix -Text $OriginalName -Length 6
        $reserve = 1 + $suffix.Length + $ext.Length
        $maxBaseLen = $roomForName - $reserve
        if ($maxBaseLen -le 0) {
            $suffix = Get-HashSuffix -Text $OriginalName -Length 10
            return "f_$suffix$ext"
        }

        $truncatedBase = if ($base.Length -gt $maxBaseLen) { $base.Substring(0, $maxBaseLen) } else { $base }
        $newName = "$truncatedBase`_$suffix$ext"

        if ((Join-Path -Path $FolderPath -ChildPath $newName).Length -gt $MaxPathLength) {
            $suffix = Get-HashSuffix -Text $OriginalName -Length 10
            $newName = "f_$suffix$ext"
        }
        return $newName
    }

    function Read-SpoFolder {
        param(
            [Parameter(Mandatory)] [Microsoft.SharePoint.Client.Folder] $Folder,
            [Parameter(Mandatory)] [string] $DestinationRoot,
            [Parameter(Mandatory)] $ActiveConnection,
            [Parameter()] [string] $LocalBaseServerRelativeUrl   # CHANGE alpha11/alpha12/alpha13/alpha14: optional trim base to avoid mirroring the full SharePoint path locally
        )
        # CHANGE alpha11/alpha12/alpha13/alpha14: compute a compact local path relative to an optional base path
        $folderServerRelativeUrl = $Folder.ServerRelativeUrl.TrimEnd('/')
        $effectiveBase = if ($PSBoundParameters.ContainsKey('LocalBaseServerRelativeUrl') -and -not [string]::IsNullOrWhiteSpace($LocalBaseServerRelativeUrl)) {
            $LocalBaseServerRelativeUrl.TrimEnd('/')
        } else {
            $Folder.Context.Web.ServerRelativeUrl.TrimEnd('/')
        }

        $relativeUrl = if ($folderServerRelativeUrl.StartsWith($effectiveBase,[System.StringComparison]::OrdinalIgnoreCase)) {
            $folderServerRelativeUrl.Substring($effectiveBase.Length).TrimStart('/')
        } else {
            $Folder.ServerRelativeUrl.Substring($Folder.Context.Web.ServerRelativeUrl.Length).TrimStart('/')
        }

        $LocalFolder = if ([string]::IsNullOrWhiteSpace($relativeUrl)) {
            $DestinationRoot
        } else {
            Join-Path -Path $DestinationRoot -ChildPath ($relativeUrl -replace "/","\")
        }
        if (!(Test-Path -LiteralPath $LocalFolder)) { New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null }

        $FolderURL = $Folder.ServerRelativeUrl.Substring($Folder.Context.Web.ServerRelativeUrl.Length)
        $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType File -Connection $ActiveConnection
        foreach ($File in $Files) {
            $SafeFileName = Get-SafeFileName -FolderPath $LocalFolder -OriginalName $File.Name -MaxPathLength 255
            $LocalFilePath = Join-Path -Path $LocalFolder -ChildPath $SafeFileName
            if ($LocalFilePath.Length -gt 255) {
                Write-log -Message "Destination path too long ($($LocalFilePath.Length)). Cannot save '$($File.Name)' to '$LocalFolder' even after renaming." -Severity ERROR -ToConsole
                continue
            }
            if ($SafeFileName -ne $File.Name) {
                Write-log -Message "Path >255 chars detected. Renaming on download: '$($File.Name)' -> '$SafeFileName'." -Severity WARN -ToConsole
            }
            if (-not (Test-Path -LiteralPath $LocalFilePath)) {
                try {
                    Get-PnPFile -ServerRelativeUrl $File.ServerRelativeUrl `
                                -Path $LocalFolder `
                                -FileName $SafeFileName `
                                -AsFile -Force `
                                -Connection $ActiveConnection
                } catch {
                    $script:errorOccurred = $true
                    Write-log -Message "Ran into an issue: $($_.ToString())" -Severity ERROR -ToConsole
                }
            }
        }

        $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType Folder -Connection $ActiveConnection
        foreach ($Sub in $SubFolders | Where-Object { $_.Name -ne "Forms" }) {
            Read-SpoFolder -Folder $Sub -DestinationRoot $DestinationRoot -ActiveConnection $ActiveConnection -LocalBaseServerRelativeUrl $effectiveBase
        }
    }

    function Remove-PnPFolderRecursive
    {
         param(
            [Parameter(Mandatory)] [Microsoft.SharePoint.Client.Folder] $Folder,
            [Parameter(Mandatory)] $ActiveConnection
        )
        $Web = Get-PnPWeb -Connection $ActiveConnection
        $FolderSiteRelativeURL = if ($Web.ServerRelativeUrl -eq "/") {
            $Folder.ServerRelativeUrl
        } else {
            $Folder.ServerRelativeUrl.Replace($Web.ServerRelativeUrl,[string]::Empty)
        }

        # files
        $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType File -Connection $ActiveConnection
        foreach ($File in $Files)
        {
         Remove-PnPFile -ServerRelativeUrl $File.ServerRelativeURL -Force -Connection $ActiveConnection
        }

        # subfolders
        $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder -Connection $ActiveConnection
        foreach ($SubFolder in $SubFolders) {
            if (($SubFolder.Name -ne "Forms") -and -not ($SubFolder.Name.StartsWith("_"))) 
            {
                Remove-PnPFolderRecursive -Folder $SubFolder -ActiveConnection $ActiveConnection
                $ParentFolderURL = $FolderSiteRelativeURL.TrimStart("/")
                    Remove-PnPFolder -Name $SubFolder.Name -Folder $ParentFolderURL -Force -Connection $ActiveConnection
            }
        }
    }
    
    function Get-RelativePath {
        param(
            [Parameter(Mandatory)] [string] $BasePath,
            [Parameter(Mandatory)] [string] $FullPath
        )
        $base = $BasePath.TrimEnd('\\')
        if ($FullPath.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $FullPath.Substring($base.Length).TrimStart('\\')
        }
        return $FullPath
    }

    function Get-FileInventory {
        param(
            [Parameter(Mandatory)] [string] $Root
        )
        $map = @{}
        if (-not (Test-Path -LiteralPath $Root)) { return $map }

        $files = Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue
        foreach ($f in $files) {
            $rel = Get-RelativePath -BasePath $Root -FullPath $f.FullName
            $rel = $rel -replace '/', '\\'
            $map[$rel] = [pscustomobject]@{
                RelPath = $rel
                FullName = $f.FullName
                LastWriteTimeUtc = $f.LastWriteTimeUtc
                Length = $f.Length
            }
        }
        return $map
    }

    function New-UniqueRelPath {
        param(
            [Parameter(Mandatory)] [string] $RelPath,
            [Parameter(Mandatory)] [hashtable] $ExistingRelMap,
            [Parameter(Mandatory)] [hashtable] $PlannedRelMap
        )
        $dir  = Split-Path $RelPath -Parent
        $leaf = Split-Path $RelPath -Leaf
        $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
        $ext  = [System.IO.Path]::GetExtension($leaf)

        $i = 1
        while ($true) {
            $newLeaf = "{0}_DUP{1}{2}" -f $base, $i, $ext
            $cand = if ([string]::IsNullOrEmpty($dir)) { $newLeaf } else { Join-Path $dir $newLeaf }
            $cand = $cand -replace '/', '\\'
            if (-not $ExistingRelMap.ContainsKey($cand) -and -not $PlannedRelMap.ContainsKey($cand)) {
                return $cand
            }
            $i++
        }
    }

    function New-MergePlan {
        param(
            [Parameter(Mandatory)] [hashtable] $SourceMap,
            [Parameter(Mandatory)] [hashtable] $TargetMap,
            [Parameter(Mandatory)] [ValidateSet('LatestWins','Rename','Fail')] [string] $Mode
        )

     #It compares two hashtables:
     #$SourceMap = files you want to merge in
     #$TargetMap = files that already exist in the destination
     #$Mode = what to do when the same relative path exists in both
     #It returns a list of planned actions ($plan) describing what should happen for each source file.

        $plan = New-Object System.Collections.Generic.List[object]
        $plannedDest = @{}

        foreach ($rel in $SourceMap.Keys) {
            $src = $SourceMap[$rel]

            if (-not $TargetMap.ContainsKey($rel)) {
                $plan.Add([pscustomobject]@{ Action='Copy'; RelPath=$rel; Src=$src.FullName; Reason='New' })
                $plannedDest[$rel] = $true
                continue
            }

            $tgt = $TargetMap[$rel]

            switch ($Mode) {
                'Fail' {
                    $plan.Add([pscustomobject]@{ Action='ConflictFail'; RelPath=$rel; Src=$src.FullName; Reason='Conflict' })
                }
                'Rename' {
                    $newRel = New-UniqueRelPath -RelPath $rel -ExistingRelMap $TargetMap -PlannedRelMap $plannedDest
                    $plan.Add([pscustomobject]@{ Action='RenameCopy'; RelPath=$newRel; Src=$src.FullName; Reason=("Conflict: kept both (renamed from {0})" -f $rel) })
                    $plannedDest[$newRel] = $true
                }
                default { # LatestWins
                    if ($src.LastWriteTimeUtc -gt $tgt.LastWriteTimeUtc) {
                        $plan.Add([pscustomobject]@{ Action='Overwrite'; RelPath=$rel; Src=$src.FullName; Reason=("Conflict: source newer ({0:o} > {1:o})" -f $src.LastWriteTimeUtc,$tgt.LastWriteTimeUtc) })
                        $plannedDest[$rel] = $true
                        $TargetMap[$rel] = $src
                    } elseif ($src.LastWriteTimeUtc -eq $tgt.LastWriteTimeUtc) {
                        $plan.Add([pscustomobject]@{ Action='Skip'; RelPath=$rel; Src=$src.FullName; Reason=("Conflict: same timestamp ({0:o}) -> keep existing" -f $src.LastWriteTimeUtc) })
                    } else {
                        $plan.Add([pscustomobject]@{ Action='Skip'; RelPath=$rel; Src=$src.FullName; Reason=("Conflict: target newer ({0:o} < {1:o}) -> keep existing" -f $src.LastWriteTimeUtc,$tgt.LastWriteTimeUtc) })
                    }
                }
            }
        }

        return $plan
    }

    function Invoke-MergePlan {
        param(
            [Parameter(Mandatory)] [System.Collections.Generic.List[object]] $Plan,
            [Parameter(Mandatory)] [ValidateSet('LatestWins','Rename','Fail')] [string] $Mode,
            [Parameter(Mandatory)] [string] $TargetRoot
        )

        if ($Mode -eq 'Fail') {
            $fail = $Plan | Where-Object { $_.Action -eq 'ConflictFail' }
            if ($fail.Count -gt 0) {
                $sample = ($fail | Select-Object -First 10 | ForEach-Object { $_.RelPath }) -join '; '
                throw "Merge aborted (ConflictMode=Fail). Conflicts detected: $sample"
            }
        }

        $stats = [ordered]@{ Copied=0; Overwritten=0; Renamed=0; Skipped=0; Conflicts=0 }

        foreach ($item in $Plan) {
            $destPath = Join-Path $TargetRoot $item.RelPath
            $destDir  = Split-Path $destPath -Parent

            switch ($item.Action) {
                'Copy' {
                    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    Copy-Item -LiteralPath $item.Src -Destination $destPath -Force
                    $stats.Copied++
                }
                'Overwrite' {
                    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    Copy-Item -LiteralPath $item.Src -Destination $destPath -Force
                    $stats.Overwritten++
                }
                'RenameCopy' {
                    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    Copy-Item -LiteralPath $item.Src -Destination $destPath -Force
                    $stats.Renamed++
                }
                'Skip' {
                    $stats.Skipped++
                }
                default { }
            }

            if ($item.Action -in @('Overwrite','RenameCopy','Skip','ConflictFail')) { $stats.Conflicts++ }
        }

        return [pscustomobject]$stats
    }

    function Merge-LocalFolderContents {

        param(
            [Parameter(Mandatory)] [string] $SourceRoot,
            [Parameter(Mandatory)] [string] $TargetRoot,
            [Parameter(Mandatory)] [ValidateSet('LatestWins','Rename','Fail')] [string] $Mode,
            [Parameter()] [string] $LogPrefix = ''
            )

        Write-log ("{0}Building in-memory inventories (Mode={1})" -f $LogPrefix,$Mode) -Sub -ToConsole -Severity INFO
        $srcMap = Get-FileInventory -Root $SourceRoot
        $tgtMap = Get-FileInventory -Root $TargetRoot

        Write-log ("{0}Source files: {1} ; Target files: {2}" -f $LogPrefix,$srcMap.Count,$tgtMap.Count) -Sub -ToConsole -Severity INFO

        $plan = New-MergePlan -SourceMap $srcMap -TargetMap $tgtMap -Mode $Mode

        $conf = $plan | Where-Object { $_.Action -in @('Overwrite','RenameCopy','Skip','ConflictFail') }
        foreach ($c in $conf) {
            Write-log ("{0}{1}: {2} -> {3}" -f $LogPrefix,$c.Action,$c.RelPath,$c.Reason) -Sub -ToConsole -Severity WARN
        }

        $stats = Invoke-MergePlan -Plan $plan -Mode $Mode -TargetRoot $TargetRoot

        Write-log ("{0}Merge result: Copied={1}, Overwritten={2}, Renamed={3}, Skipped={4}, Conflicts={5}" -f $LogPrefix,$stats.Copied,$stats.Overwritten,$stats.Renamed,$stats.Skipped,$stats.Conflicts) -Sub -ToConsole -Severity INFO
        return $stats
    }
    
    function New-CompactWorkPaths {
        param(
            [Parameter(Mandatory)] [string] $RunTemp,
            [Parameter()] [string] $RefNumber,
            [Parameter(Mandatory)] [string] $SeedText
        )
        # CHANGE alpha11/alpha12/alpha13/alpha14: ultra-short local work tree to reduce path-too-long issues during complex merges
        $refPart  = if ([string]::IsNullOrWhiteSpace($RefNumber)) { 'NOREF' } else { $RefNumber }
        $hashPart = Get-HashSuffix -Text $SeedText -Length 4
        $jobRoot  = Join-Path -Path $RunTemp -ChildPath ("W_{0}_{1}" -f $refPart,$hashPart)
        return [pscustomobject]@{
            JobRoot    = $jobRoot
            SourceRoot = (Join-Path -Path $jobRoot -ChildPath 'S')
            ZipRoot    = (Join-Path -Path $jobRoot -ChildPath 'Z')
            MergeRoot  = (Join-Path -Path $jobRoot -ChildPath 'M')
        }
    }

    function Select-DuplicateZipOutputName {
        param(
            [Parameter(Mandatory)] [string] $SourceFolderName,
            [Parameter(Mandatory)] [object[]] $MatchingZips,
            [switch] $NonInteractive
        )
        # CHANGE alpha11/alpha12/alpha13/alpha14: allow user to pick the final merged ZIP name when several ZIPs already exist for the same reference
        $candidates = New-Object System.Collections.Generic.List[string]
        $sourceFolderZipName = if ($SourceFolderName -match '(?i)\.zip$') { $SourceFolderName } else { "$SourceFolderName.zip" }
        $candidates.Add($sourceFolderZipName) | Out-Null
        foreach ($z in $MatchingZips) { if (-not $candidates.Contains($z.Name)) { $candidates.Add($z.Name) | Out-Null } }
        if ($NonInteractive) {
            Write-log ("Duplicate ZIP merge running non-interactively. Default final ZIP name: {0}" -f $sourceFolderZipName) -Severity WARN -ToConsole -Sub
            return $sourceFolderZipName
        }
        Write-Host ''
        Write-Host ("Duplicate ZIPs detected for reference in source folder '{0}'. Choose the FINAL merged ZIP name:" -f $SourceFolderName) -ForegroundColor Cyan
        Write-Host ("  1 = use the current non-compressed folder name: {0}" -f $sourceFolderZipName) -ForegroundColor Gray
        for ($i = 1; $i -lt $candidates.Count; $i++) { Write-Host (("  {0} = use existing ZIP name: {1}" -f ($i + 1), $candidates[$i])) -ForegroundColor Gray }
        $selection = Read-Host ("Enter a number between 1 and {0}. Press Enter for default (1)" -f $candidates.Count)
        if ([string]::IsNullOrWhiteSpace($selection)) { $selection = '1' }
        if (($selection -notmatch '^\d+$') -or ([int]$selection -lt 1) -or ([int]$selection -gt $candidates.Count)) {
            Write-log ("Invalid duplicate ZIP selection '{0}'. Defaulting final ZIP name to {1}" -f $selection,$sourceFolderZipName) -Severity WARN -ToConsole -Sub
            return $sourceFolderZipName
        }
        $chosenName = $candidates[[int]$selection - 1]
        Write-log ("Final merged ZIP name selected: {0}" -f $chosenName) -Severity INFO -ToConsole -Sub
        return $chosenName
    }

    function Get-NormalizedContainerInfo {
        param(
            [Parameter(Mandatory)] [string] $Path,
            [Parameter()] [string] $PreferredContainerName
        )
        # CHANGE alpha14: normalize extracted/source trees so merged ZIP keeps one consistent top-level folder
        $resolved = (Resolve-Path -LiteralPath $Path).Path
        $topItems = @(Get-ChildItem -LiteralPath $resolved -Force -ErrorAction SilentlyContinue)
        $topDirs  = @($topItems | Where-Object { $_.PSIsContainer })
        $topFiles = @($topItems | Where-Object { -not $_.PSIsContainer })

        if ($topDirs.Count -eq 1 -and $topFiles.Count -eq 0) {
            return [pscustomobject]@{
                ContentRoot    = $topDirs[0].FullName
                ContainerName  = $topDirs[0].Name
                AlreadyWrapped = $true
            }
        }

        $container = if ([string]::IsNullOrWhiteSpace($PreferredContainerName)) {
            Split-Path -Leaf $resolved
        } else {
            $PreferredContainerName
        }
        return [pscustomobject]@{
            ContentRoot    = $resolved
            ContainerName  = $container
            AlreadyWrapped = $false
        }
    }

    function Get-PnPItemServerRelativeUrl {
        param(
            [Parameter(Mandatory)] $Item,
            [Parameter(Mandatory)] [string] $FallbackUrl
        )

        foreach ($propName in @('ServerRelativeUrl','ServerRelativeURL')) {
            if ($Item.PSObject.Properties.Match($propName).Count -gt 0) {
                $value = [string]$Item.PSObject.Properties[$propName].Value
                if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
            }
        }
        return $FallbackUrl
    }

    function Get-DuplicateZipReferenceRows {
        param(
            [Parameter(Mandatory)] [object[]] $RootFolders,
            [Parameter(Mandatory)] [string] $FolderSiteRelativeURL,
            [Parameter(Mandatory)] [string] $FolderServerRelativeURL,
            [Parameter(Mandatory)] $Connection,
            [Parameter(Mandatory)] [string] $LogFolderName
        )

        Write-Host "Scanning range folders for duplicate ZIP references (per folder, consolidated output)..." -ForegroundColor Cyan

        $scanFolders = @($RootFolders | Where-Object { $_ -and $_.Name -and $_.Name -ne $LogFolderName })
        $rangeLike = @($scanFolders | Where-Object { $_.Name -match '^\d{6}\s*-\s*\d{6}$' })
        if ($rangeLike.Count -gt 0) { $scanFolders = $rangeLike }

        Write-log ("Option D: {0} folder(s) to scan (from rootfolds)" -f $scanFolders.Count) -ToConsole

        $dupRows = New-Object System.Collections.Generic.List[object]

        foreach ($rf in $scanFolders) {
            $rangeName = $rf.Name
            $subrooturl = "$FolderSiteRelativeURL/$rangeName"

            try {
                $filesAtRoot = @(Get-PnPFolderItem -FolderSiteRelativeUrl $subrooturl -ItemType File -Connection $Connection -ErrorAction Stop)
            }
            catch {
                Write-log ("Option D: Cannot enumerate files in {0} - {1}" -f $subrooturl, $_.Exception.Message) -Severity WARN -ToConsole
                continue
            }

            $subZipItems = @($filesAtRoot | Where-Object { $_.Name -match '(?i)\.zip$' })
            $zipIndex = @{}
            $zipNoRef = @()

            foreach ($z in $subZipItems) {
                $ref = Get-SixDigitReference -Name $z.Name
                if (-not [string]::IsNullOrWhiteSpace($ref)) {
                    if (-not $zipIndex.ContainsKey($ref)) { $zipIndex[$ref] = @() }
                    $zipIndex[$ref] += $z
                }
                else {
                    $zipNoRef += $z
                }
            }

            if ($dbg) {
                Write-log ("Option D: [{0}] ZIP inventory: {1} zip(s) at root, {2} distinct ref(s), {3} zip(s) without 6-digit ref" -f $rangeName, $subZipItems.Count, $zipIndex.Keys.Count, $zipNoRef.Count) -ToConsole
            }

            foreach ($k in $zipIndex.Keys) {
                if ($zipIndex[$k].Count -gt 1) {
                    foreach ($zz in $zipIndex[$k]) {
                        $fallback = "$FolderServerRelativeURL/$rangeName/$($zz.Name)"
                        $sr = Get-PnPItemServerRelativeUrl -Item $zz -FallbackUrl $fallback

                        $dupRows.Add([pscustomobject]@{
                            RangeFolder       = $rangeName
                            Ref               = $k
                            ZipName           = $zz.Name
                            ServerRelativeUrl = $sr
                            ZipItem           = $zz
                        }) | Out-Null
                    }
                }
            }
        }

        return @($dupRows)
    }

    function Write-DuplicateZipReferenceReport {
        param(
            [Parameter(Mandatory)] [object[]] $DuplicateRows
        )

        if ($DuplicateRows.Count -eq 0) {
            Write-Host "No duplicate ZIP references found in any scanned folder." -ForegroundColor Green
            return
        }

        $groups = @($DuplicateRows | Group-Object RangeFolder, Ref)
        Write-Host ("Duplicate ZIP references detected (groups): {0}" -f $groups.Count) -ForegroundColor Yellow

        foreach ($g in ($groups | Sort-Object Name)) {
            $folder = $g.Group[0].RangeFolder
            $ref    = $g.Group[0].Ref
            Write-log ("WARNING: {0} ZIPs found for ref {1} in folder {2}:" -f $g.Count, $ref, $folder) -ToConsole -Severity WARN
            foreach ($row in ($g.Group | Sort-Object ZipName)) {
                Write-log (" - {0}" -f $row.ZipName) -ToConsole -Severity WARN
            }
        }
    }

    function Export-DuplicateZipReferenceCsv {
        param(
            [Parameter(Mandatory)] [object[]] $DuplicateRows,
            [Parameter(Mandatory)] [string] $RunTemp
        )

        if ($DuplicateRows.Count -eq 0) { return }

        try {
            if (-not (Test-Path -LiteralPath $RunTemp)) { New-Item -ItemType Directory -Path $RunTemp -Force | Out-Null }
            $csvPath = Join-Path $RunTemp ("DuplicateZipRefs_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $DuplicateRows |
                Select-Object RangeFolder, Ref, ZipName, ServerRelativeUrl |
                Sort-Object RangeFolder, Ref, ZipName |
                Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host ("CSV exported to: {0}" -f $csvPath) -ForegroundColor Cyan
        }
        catch {
            Write-log "Option D: Could not export CSV - $($_.Exception.Message)" -Severity WARN -ToConsole
        }
    }

    function Select-DuplicateZipMergeOutputName {
        param(
            [Parameter(Mandatory)] [string] $RangeFolder,
            [Parameter(Mandatory)] [string] $RefNumber,
            [Parameter(Mandatory)] [object[]] $MatchingZips,
            [switch] $NonInteractive
        )

        $candidates = New-Object System.Collections.Generic.List[string]
        foreach ($z in ($MatchingZips | Sort-Object Name)) {
            if (-not $candidates.Contains($z.Name)) { $candidates.Add($z.Name) | Out-Null }
        }

        if ($candidates.Count -eq 0) { throw "No ZIP candidate available for duplicate merge." }

        if ($NonInteractive) {
            Write-log ("Duplicate ZIP merge running non-interactively. Default final ZIP name: {0}" -f $candidates[0]) -Severity WARN -ToConsole
            return $candidates[0]
        }

        Write-Host ''
        Write-Host ("Duplicate ZIP merge for folder '{0}', reference {1}. Choose the FINAL merged ZIP name:" -f $RangeFolder,$RefNumber) -ForegroundColor Cyan
        Write-Host "Only existing duplicate ZIP names are proposed, to avoid overwriting an unrelated destination file." -ForegroundColor Gray
        for ($i = 0; $i -lt $candidates.Count; $i++) {
            Write-Host ("  {0} = {1}" -f ($i + 1), $candidates[$i]) -ForegroundColor Gray
        }

        $selection = Read-Host ("Enter a number between 1 and {0}. Press Enter for default (1). Enter S to skip this group" -f $candidates.Count)
        if ([string]::IsNullOrWhiteSpace($selection)) { $selection = '1' }
        if ($selection -match '^[sS]$') { return $null }

        if (($selection -notmatch '^\d+$') -or ([int]$selection -lt 1) -or ([int]$selection -gt $candidates.Count)) {
            Write-log ("Invalid duplicate ZIP merge selection '{0}'. Skipping ref {1} in folder {2}." -f $selection,$RefNumber,$RangeFolder) -Severity WARN -ToConsole
            return $null
        }

        $chosenName = $candidates[[int]$selection - 1]
        Write-log ("Final merged ZIP name selected for ref {0} in folder {1}: {2}" -f $RefNumber,$RangeFolder,$chosenName) -Severity INFO -ToConsole
        return $chosenName
    }

    function Test-DestinationZipUploadSafe {
        param(
            [Parameter(Mandatory)] [string] $FolderSiteRelativeUrl,
            [Parameter(Mandatory)] [string] $ServerRelativeSubRoot,
            [Parameter(Mandatory)] [string] $ZipFileName,
            [Parameter(Mandatory)] [string] $RefNumber,
            [Parameter()] [string[]] $ExpectedExistingZipNames = @(),
            [Parameter(Mandatory)] $Connection,
            [switch] $Sub
        )

        $expected = @($ExpectedExistingZipNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        try {
            $filesAtRoot = @(Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeUrl -ItemType File -Connection $Connection -ErrorAction Stop)
        }
        catch {
            Write-log ("[REF {0}] Could not re-check destination ZIP inventory before upload to {1}. Skipping upload. Error: {2}" -f $RefNumber,$FolderSiteRelativeUrl,$_.Exception.Message) -Severity WARN -ToConsole -Sub:$Sub
            return $false
        }

        $zipItems = @($filesAtRoot | Where-Object { $_.Name -match '(?i)\.zip$' })
        $exactNameCollisions = @($zipItems | Where-Object { $_.Name -eq $ZipFileName -and ($expected -notcontains $_.Name) })
        if ($exactNameCollisions.Count -gt 0) {
            foreach ($collision in $exactNameCollisions) {
                Write-log ("[REF {0}] SKIP upload: destination ZIP name already exists and was not part of the planned merge/update set: {1}" -f $RefNumber,$collision.Name) -Severity WARN -ToConsole -Sub:$Sub
            }
            return $false
        }

        $sameRefUnexpected = @($zipItems | Where-Object {
            $existingRef = Get-SixDigitReference -Name $_.Name
            ($existingRef -eq $RefNumber) -and ($expected -notcontains $_.Name)
        })

        if ($sameRefUnexpected.Count -gt 0) {
            Write-log ("[REF {0}] SKIP upload: unexpected destination ZIP(s) with the same reference already exist. No overwrite/merge/delete will be attempted." -f $RefNumber) -Severity WARN -ToConsole -Sub:$Sub
            foreach ($collision in ($sameRefUnexpected | Sort-Object Name)) {
                $fallback = "$ServerRelativeSubRoot/$($collision.Name)"
                $sr = Get-PnPItemServerRelativeUrl -Item $collision -FallbackUrl $fallback
                Write-log ("[REF {0}] Blocking destination ZIP: {1} | {2}" -f $RefNumber,$collision.Name,$sr) -Severity WARN -ToConsole -Sub:$Sub
            }
            return $false
        }

        return $true
    }

    function Invoke-DuplicateZipGroupMerge {
        param(
            [Parameter(Mandatory)] [object[]] $DuplicateRows,
            [Parameter(Mandatory)] [string] $FolderSiteRelativeURL,
            [Parameter(Mandatory)] [string] $FolderServerRelativeURL,
            [Parameter(Mandatory)] [string] $RunTemp,
            [Parameter(Mandatory)] $Connection,
            [Parameter(Mandatory)] [ValidateSet('LatestWins','Rename','Fail')] [string] $ConflictMode,
            [switch] $NonInteractive
        )

        if ($DuplicateRows.Count -lt 2) {
            return [pscustomobject]@{ Status='Skipped'; Reason='Less than two ZIPs'; RangeFolder=''; Ref=''; FinalZip='' }
        }

        $rangeName = [string]$DuplicateRows[0].RangeFolder
        $refNumber = [string]$DuplicateRows[0].Ref
        $subrooturl = "$FolderSiteRelativeURL/$rangeName"
        $serverRelativeSubRoot = "$FolderServerRelativeURL/$rangeName"
        $matchingZips = @($DuplicateRows | Sort-Object ZipName | ForEach-Object { $_.ZipItem })
        $expectedZipNames = @($matchingZips | ForEach-Object { $_.Name })

        Write-log ("Option D merge: processing {0} duplicate ZIP(s) for ref {1} in folder {2}." -f $matchingZips.Count,$refNumber,$rangeName) -Severity WARN -ToConsole

        $canonicalZipName = Select-DuplicateZipMergeOutputName -RangeFolder $rangeName -RefNumber $refNumber -MatchingZips $matchingZips -NonInteractive:$NonInteractive
        if ([string]::IsNullOrWhiteSpace($canonicalZipName)) {
            Write-log ("Option D merge: skipped by user for ref {0} in folder {1}." -f $refNumber,$rangeName) -Severity WARN -ToConsole
            return [pscustomobject]@{ Status='Skipped'; Reason='Skipped by user'; RangeFolder=$rangeName; Ref=$refNumber; FinalZip='' }
        }

        if (-not (Test-DestinationZipUploadSafe -FolderSiteRelativeUrl $subrooturl -ServerRelativeSubRoot $serverRelativeSubRoot -ZipFileName $canonicalZipName -RefNumber $refNumber -ExpectedExistingZipNames $expectedZipNames -Connection $Connection)) {
            Write-log ("Option D merge: skipped ref {0} in folder {1} because destination changed or contains an unexpected ZIP collision." -f $refNumber,$rangeName) -Severity WARN -ToConsole
            return [pscustomobject]@{ Status='Skipped'; Reason='Destination collision'; RangeFolder=$rangeName; Ref=$refNumber; FinalZip=$canonicalZipName }
        }

        $workPaths = New-CompactWorkPaths -RunTemp $RunTemp -RefNumber $refNumber -SeedText ("OptionD|{0}|{1}" -f $rangeName,$refNumber)
        foreach ($p in @($workPaths.JobRoot,$workPaths.SourceRoot,$workPaths.ZipRoot,$workPaths.MergeRoot)) {
            if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }

        $mergeStage    = $workPaths.MergeRoot
        $existingZipDl = Join-Path -Path $workPaths.JobRoot -ChildPath 'ZE'
        foreach ($p in @($mergeStage,$existingZipDl)) {
            if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }

        $canonicalContainerName = [System.IO.Path]::GetFileNameWithoutExtension($canonicalZipName)
        $canonicalRoot = Join-Path -Path $mergeStage -ChildPath $canonicalContainerName
        if (-not (Test-Path -LiteralPath $canonicalRoot)) { New-Item -ItemType Directory -Path $canonicalRoot -Force | Out-Null }

        $ziperr = 0
        $zipCounter = 0

        foreach ($z in $matchingZips) {
            $zipCounter++
            $localZip = Join-Path -Path $existingZipDl -ChildPath $z.Name
            $zipExtract = Join-Path -Path $existingZipDl -ChildPath ("X{0:D2}" -f $zipCounter)

            try {
                Write-log ("Option D merge: downloading ZIP {0}" -f $z.Name) -Severity INFO -ToConsole
                Get-PnPFile -ServerRelativeUrl "$serverRelativeSubRoot/$($z.Name)" `
                            -Path $existingZipDl `
                            -FileName $($z.Name) `
                            -AsFile -Force `
                            -Connection $Connection

                if (Test-Path -LiteralPath $zipExtract) { Remove-Item -LiteralPath $zipExtract -Recurse -Force -ErrorAction SilentlyContinue }
                New-Item -ItemType Directory -Path $zipExtract -Force | Out-Null

                Write-log ("Option D merge: extracting {0}" -f $z.Name) -Severity INFO -ToConsole
                Expand-Archive -LiteralPath $localZip -DestinationPath $zipExtract -Force

                $zipInfo = Get-NormalizedContainerInfo -Path $zipExtract -PreferredContainerName $canonicalContainerName
                Merge-LocalFolderContents -SourceRoot $zipInfo.ContentRoot -TargetRoot $canonicalRoot -Mode $ConflictMode -LogPrefix ("[OptionD ZIP:{0}] " -f $z.Name) | Out-Null
            }
            catch {
                Write-log ("Option D merge: failed to download/extract/merge {0}: {1}" -f $z.Name,$_.Exception.Message) -Severity ERROR -ToConsole
                $ziperr++
            }
            finally {
                if (Test-Path -LiteralPath $zipExtract) { Remove-Item -LiteralPath $zipExtract -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }

        $zipdestination = Join-Path -Path $workPaths.ZipRoot -ChildPath $canonicalZipName

        if ($ziperr -eq 0) {
            try {
                if (Test-Path -LiteralPath $zipdestination) { Remove-Item -LiteralPath $zipdestination -Force -ErrorAction SilentlyContinue }
                Write-log ("Option D merge: creating merged ZIP {0}" -f $canonicalZipName) -Severity INFO -ToConsole
                Compress-Archive -Path (Join-Path $mergeStage '*') -DestinationPath $zipdestination -Force

                $mergedCount = $null
                $ZipFile = $null
                try {
                    $ZipFile = [System.IO.Compression.ZipFile]::Open($zipdestination,[System.IO.Compression.ZipArchiveMode]::Read)
                    $mergedCount = $ZipFile.Entries.Count
                }
                finally {
                    if ($ZipFile) { $ZipFile.Dispose() }
                }
                Write-log ("Option D merge: merged ZIP created: {0} now holds {1} files" -f $canonicalZipName,$mergedCount) -Severity INFO -ToConsole
            }
            catch {
                Write-log ("Option D merge: failed to create merged ZIP for ref {0}: {1}" -f $refNumber,$_.Exception.Message) -Severity ERROR -ToConsole
                $ziperr++
            }
        }

        if ($ziperr -eq 0) {
            try {
                $zipServerRelativeUrl = "$serverRelativeSubRoot/$canonicalZipName"
                Write-log ("Option D merge: uploading final ZIP {0} to {1}" -f $canonicalZipName,$serverRelativeSubRoot) -Severity INFO -ToConsole
                try { Set-PnPFileCheckedOut -Url $zipServerRelativeUrl -Connection $Connection -ErrorAction SilentlyContinue } catch { }
                Add-PnPFile -Path $zipdestination -Folder ($subrooturl.TrimStart("/")) -Connection $Connection | Out-Null
                try { Set-PnPFileCheckedIn -Url $zipServerRelativeUrl -CheckinType MajorCheckIn -Connection $Connection -ErrorAction SilentlyContinue } catch { }
            }
            catch {
                Write-log ("Option D merge: failed to upload final ZIP {0}: {1}" -f $canonicalZipName,$_.Exception.Message) -Severity ERROR -ToConsole
                $ziperr++
            }
        }

        if ($ziperr -eq 0) {
            $dupes = @($matchingZips | Where-Object { $_.Name -ne $canonicalZipName })
            foreach ($d in $dupes) {
                Write-log ("Option D merge: deleting duplicate ZIP on SharePoint after successful merge: {0}" -f $d.Name) -Severity WARN -ToConsole
                    try {
                        Remove-PnPFile -ServerRelativeUrl "$serverRelativeSubRoot/$($d.Name)" -Force -Connection $Connection
                    }
                    catch {
                        Write-log ("Option D merge: failed to delete duplicate ZIP {0}: {1}" -f $d.Name,$_.Exception.Message) -Severity WARN -ToConsole
                    }
            }
        }

        if (Test-Path -LiteralPath $mergeStage) { Remove-Item -LiteralPath $mergeStage -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $existingZipDl) { Remove-Item -LiteralPath $existingZipDl -Recurse -Force -ErrorAction SilentlyContinue }

        if ($ziperr -eq 0) {
            return [pscustomobject]@{ Status='Merged'; Reason='OK'; RangeFolder=$rangeName; Ref=$refNumber; FinalZip=$canonicalZipName }
        }
        return [pscustomobject]@{ Status='Error'; Reason='Merge/upload error'; RangeFolder=$rangeName; Ref=$refNumber; FinalZip=$canonicalZipName }
    }

    function Invoke-DuplicateZipMergeFromRows {
        param(
            [Parameter(Mandatory)] [object[]] $DuplicateRows,
            [Parameter(Mandatory)] [string] $FolderSiteRelativeURL,
            [Parameter(Mandatory)] [string] $FolderServerRelativeURL,
            [Parameter(Mandatory)] [string] $RunTemp,
            [Parameter(Mandatory)] $Connection,
            [Parameter(Mandatory)] [ValidateSet('LatestWins','Rename','Fail')] [string] $ConflictMode,
            [Parameter(Mandatory)] [ValidateSet('SelectedRoot','AllRoots')] [string] $Scope,
            [switch] $NonInteractive
        )

        if ($DuplicateRows.Count -eq 0) { return @() }

        $rowsToMerge = @($DuplicateRows)

        if ($Scope -eq 'SelectedRoot') {
            $rootGroups = @($DuplicateRows | Group-Object RangeFolder | Sort-Object Name)
            if ($rootGroups.Count -eq 1) {
                $rowsToMerge = @($rootGroups[0].Group)
                Write-log ("Option D merge: only one root folder has duplicates; selected {0}." -f $rootGroups[0].Name) -Severity INFO -ToConsole
            }
            else {
                Write-Host ''
                Write-Host "Root folders with duplicate ZIP references:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $rootGroups.Count; $i++) {
                    $groupCount = @($rootGroups[$i].Group | Group-Object RangeFolder, Ref).Count
                    Write-Host ("  {0} = {1} ({2} duplicate reference group(s))" -f ($i + 1), $rootGroups[$i].Name, $groupCount)
                }
                $selection = Read-Host ("Enter the root folder number to merge. Press Enter to cancel")
                if ([string]::IsNullOrWhiteSpace($selection)) {
                    Write-log "Option D merge: cancelled before selecting a root folder." -Severity WARN -ToConsole
                    return @()
                }
                if (($selection -notmatch '^\d+$') -or ([int]$selection -lt 1) -or ([int]$selection -gt $rootGroups.Count)) {
                    Write-log ("Option D merge: invalid root folder selection '{0}'. Cancelled." -f $selection) -Severity WARN -ToConsole
                    return @()
                }
                $rowsToMerge = @($rootGroups[[int]$selection - 1].Group)
            }
        }

        $mergeResults = New-Object System.Collections.Generic.List[object]
        $groups = @($rowsToMerge | Group-Object RangeFolder, Ref | Sort-Object Name)
        Write-log ("Option D merge: {0} duplicate reference group(s) selected for merge." -f $groups.Count) -Severity WARN -ToConsole

        foreach ($g in $groups) {
            $result = Invoke-DuplicateZipGroupMerge -DuplicateRows @($g.Group) `
                                                     -FolderSiteRelativeURL $FolderSiteRelativeURL `
                                                     -FolderServerRelativeURL $FolderServerRelativeURL `
                                                     -RunTemp $RunTemp `
                                                     -Connection $Connection `
                                                     -ConflictMode $ConflictMode `
                                                     -NonInteractive:$NonInteractive
            $mergeResults.Add($result) | Out-Null
        }

        Write-Host ''
        Write-Host "Option D merge summary:" -ForegroundColor Cyan
        $mergeResults | Format-Table Status, RangeFolder, Ref, FinalZip, Reason -AutoSize | Out-String -Width 300 | Write-Host
        return @($mergeResults)
    }

    # ---- CONNECT ----
    Clear-Host
    Write-log "Running script version $scriptVersion" -ToConsole
    Write-log "Connecting to SharePoint site $SiteUrl" -ToConsole

    try
    {
        # Prefer modern auth, but will always fail under current ING context.
        # simply close the window and it will use the alternative legacy method which work
        try
        {
            $ActiveConnection = Connect-PnPOnline -Url $SiteUrl  -Interactive -ReturnConnection     
            
        } catch
        {
            # Fallback for ISE/old runtimes but requires pnp.powershell 1.12 module (old)
            Write-log "Falling back to legacy authentication." -Severity WARN -ToConsole
            $ActiveConnection = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection
        }
    }
    catch
    {
        Write-log "Cannot connect to target SharePoint site, execution aborted." -Severity FATAL -ToConsole
        throw
    }

    # ---- DISCOVER ROOT FOLDERS ----
    $rootfolds = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -Connection $ActiveConnection
    Write-log "Target Folder $FolderSiteRelativeURL has $($rootfolds.Count) subfolders to process" -ToConsole

    # Show a quick snapshot:
    Write-Host "Scanning for folders to process on source: $($FolderServerRelativeURL)" -ForegroundColor Cyan
    for ($i = 0; $i -lt $rootfolds.Count; $i++)
    {
        $curfol = "$FolderSiteRelativeURL/$($rootfolds[$i].Name)"
        $follist = @(Get-PnPFolderItem -FolderSiteRelativeUrl $curfol -ItemType Folder -Connection $ActiveConnection |
                   Where-Object { $_.Name -ne $LogFolderName })
        $eligibleFolCount = @($follist | Where-Object { -not [string]::IsNullOrWhiteSpace((Get-SixDigitReference -Name $_.Name)) }).Count
        $skippedFolCount  = $follist.Count - $eligibleFolCount
        Write-Host ("{0}. {1} ({2} eligible subfolders, {3} skipped: no 6-digit ref)" -f ($i+1), $rootfolds[$i].Name, $eligibleFolCount, $skippedFolCount)
    }

    # Selection (unless -ProcessAll)
    $selected = $rootfolds

    if (-not $ProcessAll)
    {
        Write-Host "A. All folders"
        Write-Host "C. Check folder/ZIP placement against 6-digit ranges (display only, no issue logging)"
        Write-Host "D. Find duplicate ZIP references"
        Write-Host "Q. Quit"
        $folderSelection = Read-Host "Enter the number of the folder to process, or 'A', 'C', 'D', 'Q' (default 'Q')"
        if ([string]::IsNullOrWhiteSpace($folderSelection)) { $folderSelection = 'Q' }
        $folderSelection = $folderSelection.ToUpperInvariant()

        # OPTION C: display-only check that uncompressed folders and already-compressed ZIP files are in the right range folder.
        # This option deliberately avoids Write-log for detected placement issues.
        if ($folderSelection -eq 'C')
        {
            Invoke-PlacementCheck -RootFolders $rootfolds `
                                  -FolderSiteRelativeURL $FolderSiteRelativeURL `
                                  -FolderServerRelativeURL $FolderServerRelativeURL `
                                  -Connection $ActiveConnection `
                                  -LogFolderName $LogFolderName
            return
        } # End of C Selection

        # OPTION D (FINAL): per-range-folder duplicate detection, then consolidated report ---
        # Requirement: $rootfolds is the list of range folders to scan (e.g. '717000 - 717999').
        # For EACH folder: scan ZIPs at that folder root, build duplicates within that folder, then move on.
        # At the end: display all duplicates across all scanned folders at once.

        if ($folderSelection -eq 'D')
        {
            $dupRows = @(Get-DuplicateZipReferenceRows -RootFolders $rootfolds `
                                                   -FolderSiteRelativeURL $FolderSiteRelativeURL `
                                                   -FolderServerRelativeURL $FolderServerRelativeURL `
                                                   -Connection $ActiveConnection `
                                                   -LogFolderName $LogFolderName)

            Write-DuplicateZipReferenceReport -DuplicateRows $dupRows
            Export-DuplicateZipReferenceCsv -DuplicateRows $dupRows -RunTemp $RunTemp

            if ($dupRows.Count -gt 0) {
                Write-Host ''
                Write-Host "Duplicate ZIP merge options:" -ForegroundColor Cyan
                Write-Host "R. Merge all duplicate ZIP groups in one selected root folder"
                Write-Host "G. Merge all duplicate ZIP groups in every scanned root folder"
                Write-Host "Q. Quit without merging"
                $mergeSelection = Read-Host "Enter 'R', 'G' or 'Q' (default 'Q')"
                if ([string]::IsNullOrWhiteSpace($mergeSelection)) { $mergeSelection = 'Q' }
                $mergeSelection = $mergeSelection.ToUpperInvariant()

                if ($mergeSelection -eq 'R') {
                    Invoke-DuplicateZipMergeFromRows -DuplicateRows $dupRows `
                                                     -FolderSiteRelativeURL $FolderSiteRelativeURL `
                                                     -FolderServerRelativeURL $FolderServerRelativeURL `
                                                     -RunTemp $RunTemp `
                                                     -Connection $ActiveConnection `
                                                     -ConflictMode $ConflictMode `
                                                     -Scope SelectedRoot | Out-Null
                }
                elseif ($mergeSelection -eq 'G') {
                    Invoke-DuplicateZipMergeFromRows -DuplicateRows $dupRows `
                                                     -FolderSiteRelativeURL $FolderSiteRelativeURL `
                                                     -FolderServerRelativeURL $FolderServerRelativeURL `
                                                     -RunTemp $RunTemp `
                                                     -Connection $ActiveConnection `
                                                     -ConflictMode $ConflictMode `
                                                     -Scope AllRoots | Out-Null
                }
                elseif ($mergeSelection -eq 'Q') {
                    Write-log "Option D: duplicate scan completed; merge skipped by user." -Severity INFO -ToConsole
                }
                else {
                    Write-log ("Option D: invalid merge selection '{0}'. Merge skipped." -f $mergeSelection) -Severity WARN -ToConsole
                }
            }

            return
        } # End of D Selection


        if ([string]::IsNullOrWhiteSpace($folderSelection)) { $folderSelection = 'A' }
        if ($folderSelection -eq 'Q')
         {
         Write-log "Script terminated by user." -Severity FATAL -ToConsole
         return
         }
        if ($folderSelection -ne 'A')
         {
          if (($folderSelection -notmatch '^\d+$') -or ([int]$folderSelection -lt 1) -or ([int]$folderSelection -gt $rootfolds.Count))
           {
            Write-log ("Invalid selection '{0}'. Script terminated." -f $folderSelection) -Severity FATAL -ToConsole
            return
           }
          $selected = @($rootfolds[[int]$folderSelection - 1])
         }
    }  # End of A Selection


     ###########################################################################################################
     #        FOR EACH FOLDER IN THE ROOT SELECTED FOLDERS (example: /Shared Documents/_Archived Files/Folder  #
     ###########################################################################################################

    foreach ($rootfold in $selected)
    {
     Write-log "Processing folder: $($rootfold.Name)" -ToConsole
     $subrooturl = "$FolderSiteRelativeURL/$($rootfold.Name)"         # site-relative
     $serverRelativeSubRoot = "$FolderServerRelativeURL/$($rootfold.Name)"  # server-relative
     
     # Ensure COMPRESSION HISTORY exists before filtering, so no-reference skips can also be audited.
     Resolve-PnPFolder -SiteRelativePath "$subrooturl/$LogFolderName" -Connection $ActiveConnection | Out-Null

     # Per-root sub-log must exist before the no-reference filter.
     # CHANGE alpha16: skipped folders are written to this file via Write-log -Sub.
     $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
     $sublogfile = "log_$($rootfold.Name)_$timestamp.txt"
     $subLogPath = Join-Path -Path $RunTemp -ChildPath $sublogfile
     "" | Out-File -FilePath $subLogPath -Encoding UTF8 -Force

     # get list of non compressed folders excluding the compression log folder
     # CHANGE alpha15/alpha16: only folders with a standalone 6-digit reference are eligible.
     # This filter runs BEFORE any local workspace creation, SharePoint folder download, ZIP merge, upload or delete.
     $allSubitems = @(Get-PnPFolderItem -FolderSiteRelativeUrl $subrooturl -ItemType Folder -Connection $ActiveConnection | Where-Object { $_.Name -ne $LogFolderName })
     $subitemsWithoutRef = @($allSubitems | Where-Object { [string]::IsNullOrWhiteSpace((Get-SixDigitReference -Name $_.Name)) })
     $subitems = @($allSubitems | Where-Object { -not [string]::IsNullOrWhiteSpace((Get-SixDigitReference -Name $_.Name)) })
     $subqty = $subitems.Count

     if ($subitemsWithoutRef.Count -gt 0)
      {
       Write-log ("Skipping {0} subfolder(s) without standalone 6-digit reference before any processing/download." -f $subitemsWithoutRef.Count) -Severity WARN -ToConsole -Sub
       foreach ($skippedSubitem in ($subitemsWithoutRef | Sort-Object Name))
        {
         $skippedSubitemUrl = "$subrooturl/$($skippedSubitem.Name)"
         Write-log ("[SKIP NO REF] Source folder skipped before download: {0}. Reason: folder name does not contain a standalone 6-digit reference. No download, ZIP update, upload or delete will be attempted." -f $skippedSubitemUrl) -Severity WARN -ToConsole -Sub
        }
      }
     
     if ($subqty -eq 0)
      {
       Write-log "This folder has no eligible subfolder to compress; folders without standalone 6-digit reference were skipped before download." -ToConsole -Sub
       if (Test-Path -LiteralPath $subLogPath)
        {
         try
          {
           $logDestSiteRel = "$subrooturl/$LogFolderName"
           $logDestFolder = Get-PnPFolder -Url $logDestSiteRel -Connection $ActiveConnection
           Write-log "Uploading compression log $sublogfile to $($logDestFolder.ServerRelativeUrl)" -Severity INFO -ToConsole
           Add-PnPFile -Path $subLogPath -Folder $logDestSiteRel.TrimStart("/") -Connection $ActiveConnection | Out-Null
          }
         catch
          {
           Write-log "Failed to upload compression log $sublogfile : $($_.Exception.Message)" -Severity WARN -ToConsole
          }
        }
       Write-Host ""
       continue
      }

     Write-log "This folder has $subqty eligible subfolder(s) to compress" -ToConsole -Sub

     # Zips directly under the subroot (rootfold) - inventory in memory
     # Build an in-memory index of ZIPs by 6-digit reference for fast lookup and better traceability
     Write-log "Getting list of existing ZIP files at $subrooturl (inventory in memory)" -ToConsole

     $filesAtRoot = Get-PnPFolderItem -FolderSiteRelativeUrl $subrooturl -ItemType File -Connection $ActiveConnection
     $subZipItems = $filesAtRoot | Where-Object { $_.Name -match '(?i)\.zip$' }

     # Index: 6-digit ref -> list of zip items (supports duplicates)
     $zipIndex = @{}
     $zipNoRef = @()

     foreach ($z in $subZipItems)
     {
            $ref = Get-SixDigitReference -Name $z.Name
            if (-not [string]::IsNullOrWhiteSpace($ref)) {
                if (-not $zipIndex.ContainsKey($ref)) { $zipIndex[$ref] = @() }
                $zipIndex[$ref] += $z
            } else {
                $zipNoRef += $z
            }
        }
     Write-log ("ZIP inventory: {0} zip(s) found at root, {1} distinct ref(s), {2} zip(s) without 6-digit ref" -f $subZipItems.Count, $zipIndex.Keys.Count, $zipNoRef.Count) -ToConsole
     
     foreach ($k in $zipIndex.Keys)
     {
            if ($zipIndex[$k].Count -gt 1) {
                Write-log ("WARNING: {0} ZIPs found for ref {1}:" -f $zipIndex[$k].Count, $k) -ToConsole -Severity WARN
                foreach ($zz in $zipIndex[$k]) {
                    Write-log (" - {0}" -f $zz.Name) -ToConsole -Severity WARN
                }
            }
        }

     ###########################################################################################################
     # FOR EACH SUBFOLDER IN THE CURRENT SELECTION OF FOLDERS (example: Folder 964000 - 964999/964145 - xx     #
     ###########################################################################################################

     foreach ($subitem in $subitems)
     {
            $subitemUrl = "$subrooturl/$($subitem.Name)"  # site-relative

            # CHANGE alpha15: mandatory guard. If this ever gets reached without a valid reference,
            # stop immediately before empty-check, workspace creation, download, ZIP handling or deletion.
            $refNumber = Get-SixDigitReference -Name $subitem.Name
            if ([string]::IsNullOrWhiteSpace($refNumber)) {
                Write-log "Skipping folder $subitemUrl because its name has no standalone 6-digit reference. No download, ZIP update, upload or delete will be attempted." -Sub -ToConsole -Severity WARN
                continue
            }

            Write-log "Processing compression of $subitemUrl folder" -Sub -ToConsole
            Write-log "Folder reference: $refNumber" -Sub -ToConsole
            Write-Host ""

            # Determine if empty (direct children)
            $itemsInFolder = (Get-PnPFolderItem -FolderSiteRelativeUrl $subitemUrl -Connection $ActiveConnection).Count
            if ($itemsInFolder -le 0) {
                Write-log "Folder $subitemUrl is empty, nothing to compress, skipping it." -Sub -ToConsole -Severity WARN
                continue
            }

            # CHANGE alpha11/alpha12/alpha13/alpha14: create a very short local work structure for this current $subitem only
            $workPaths = New-CompactWorkPaths -RunTemp $RunTemp -RefNumber $refNumber -SeedText $subitemUrl
            foreach ($p in @($workPaths.JobRoot,$workPaths.SourceRoot,$workPaths.ZipRoot,$workPaths.MergeRoot)) {
                if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
                New-Item -ItemType Directory -Path $p -Force | Out-Null
            }
            
            # $LocalFolderRoot = $workPaths.JobRoot  - VS code says variable not used anymore

            # CHANGE alpha11/alpha12/alpha13/alpha14: download only current $subitem tree under SourceRoot, not the full SharePoint path
            $FolderObj = Get-PnPFolder -Url  $subitemUrl -Connection $ActiveConnection
            Write-log "Downloading local copy of $($FolderObj.ServerRelativeUrl) into short work area $($workPaths.SourceRoot)" -Sub -ToConsole -Severity INFO

            $script:errorOccurred = $false
            Read-SpoFolder -Folder $FolderObj -DestinationRoot $workPaths.SourceRoot -ActiveConnection $ActiveConnection -LocalBaseServerRelativeUrl $serverRelativeSubRoot

            if ($script:errorOccurred) {
                Write-log "Errors occurred during download from SharePoint, skipping folder" -Sub -Severity ERROR -ToConsole
                continue
            }

            $zipsourcefolder = Join-Path -Path $workPaths.SourceRoot -ChildPath $($subitem.Name)

            # See if a matching ZIP exists under the subroot by same 6-digit ref (not exact folder name)
            # Use per-root in-memory ZIP index (zipIndex) instead of re-scanning / filtering items
            $matchingZips = @()
            if ($zipIndex.ContainsKey($refNumber)) { $matchingZips = @($zipIndex[$refNumber]) }

            $ziperr = 0

             if ($matchingZips.Count -eq 1) # ONE MATCHING ZIP, APPEND TO EXISTING ZIP
             {
                $matchingZip = $matchingZips[0]
                Write-log "Found a matching ZIP on SharePoint (memory): $($matchingZip.Name)" -Sub -ToConsole -Severity INFO

                # CHANGE alpha12/alpha13/alpha14: verified path alignment - existing ZIP is downloaded into ZipRoot and opened from the same ZipRoot path
                Write-log "Downloading ZIP from SharePoint: $subrooturl/$($matchingZip.Name) into short ZIP work area $($workPaths.ZipRoot)" -Sub -ToConsole -Severity INFO

                $zipLocalPath = Join-Path -Path $workPaths.ZipRoot -ChildPath $($matchingZip.Name)
                try {
                    Get-PnPFile -ServerRelativeUrl "$serverRelativeSubRoot/$($matchingZip.Name)" `
                                -Path $workPaths.ZipRoot `
                                -FileName $($matchingZip.Name) `
                                -AsFile -Force `
                                -Connection $ActiveConnection

                    # Open to count entries
                    $zipBeforeCount = $null
                    $ZipFile = $null
                    try {
                        $ZipFile = [System.IO.Compression.ZipFile]::Open($zipLocalPath,[System.IO.Compression.ZipArchiveMode]::Read)
                        $zipBeforeCount = $ZipFile.Entries.Count
                    } finally {
                        if ($ZipFile) { $ZipFile.Dispose() }
                    }
                    Write-log "$($matchingZip.Name) currently holds $zipBeforeCount files" -Sub -Severity INFO -ToConsole

                    #True merge using staging folder + conflict strategy ($ConflictMode)
                    $mergeStage = $workPaths.MergeRoot   # CHANGE alpha11/alpha12/alpha13/alpha14: use a short dedicated merge staging folder

                    if (Test-Path -LiteralPath $mergeStage) {
                        Remove-Item -LiteralPath $mergeStage -Recurse -Force -ErrorAction SilentlyContinue
                    }
                    New-Item -ItemType Directory -Path $mergeStage -Force | Out-Null

                    Write-log "Extracting existing ZIP into staging: $mergeStage" -Sub -ToConsole -Severity INFO
                    Expand-Archive -LiteralPath $zipLocalPath -DestinationPath $mergeStage -Force

                    # CHANGE alpha14: normalize both existing ZIP and incoming source under one canonical top-level folder
                    $existingInfo = Get-NormalizedContainerInfo -Path $mergeStage -PreferredContainerName $subitem.Name
                    $sourceInfo   = Get-NormalizedContainerInfo -Path $zipsourcefolder -PreferredContainerName $subitem.Name
                    $canonicalContainerName = $existingInfo.ContainerName
                    if ([string]::IsNullOrWhiteSpace($canonicalContainerName)) { $canonicalContainerName = $subitem.Name }
                    $canonicalRoot = if ($existingInfo.AlreadyWrapped) { Join-Path -Path $mergeStage -ChildPath $canonicalContainerName } else { Join-Path -Path $mergeStage -ChildPath $canonicalContainerName }
                    if (-not (Test-Path -LiteralPath $canonicalRoot)) { New-Item -ItemType Directory -Path $canonicalRoot -Force | Out-Null }
                    Merge-LocalFolderContents -SourceRoot $sourceInfo.ContentRoot -TargetRoot $canonicalRoot -Mode $ConflictMode -LogPrefix "[ExistingZIP Merge] " | Out-Null

                    Write-log "Rebuilding ZIP from staging (true merge) into: $zipLocalPath" -Sub -ToConsole -Severity INFO
                    if (Test-Path -LiteralPath $zipLocalPath) {
                        Remove-Item -LiteralPath $zipLocalPath -Force -ErrorAction SilentlyContinue
                    }
                    Compress-Archive -Path (Join-Path $mergeStage '*') -DestinationPath $zipLocalPath -Force

                    # Count after
                    $zipAfterCount = $null
                    try {
                        $ZipFile = [System.IO.Compression.ZipFile]::Open($zipLocalPath,[System.IO.Compression.ZipArchiveMode]::Read)
                        $zipAfterCount = $ZipFile.Entries.Count
                    } finally {
                        if ($ZipFile) { $ZipFile.Dispose() }
                    }
                    Write-log "Updated $($matchingZip.Name). File now holds $zipAfterCount files" -Sub -Severity INFO -ToConsole

                    # Cleanup staging
                    if (Test-Path -LiteralPath $mergeStage) {
                        Remove-Item -LiteralPath $mergeStage -Recurse -Force -ErrorAction SilentlyContinue
                    }

                } catch {
                    Write-log "Failed to download/update existing ZIP '$($matchingZip.Name)': $($_.Exception.Message)" -Sub -Severity ERROR -ToConsole
                    $ziperr++
                }
                $zipdestination = $zipLocalPath

             } 
             elseif ($matchingZips.Count -gt 1) # SEVERAL MATCHING ZIP, MERGE
             {
              #Apply conflict strategy also between ZIPs by extracting each ZIP separately then merging via plan
              Write-log ("Multiple ZIPs found for reference {0}. Will merge them, then add new documents." -f $refNumber) -Sub -ToConsole -Severity WARN
              
              foreach ($z in $matchingZips) {
                    Write-log (" - Candidate ZIP: {0}" -f $z.Name) -Sub -ToConsole -Severity WARN
                }
              $mergeStage    = $workPaths.MergeRoot   # CHANGE alpha11/alpha12/alpha13/alpha14: keep merge staging on a very short path
              $existingZipDl = Join-Path -Path $workPaths.JobRoot -ChildPath 'ZE'   # CHANGE alpha13/alpha14: temporary area for downloaded duplicate ZIPs; keep separate from final ZipRoot so cleanup does not delete the final ZIP

              foreach ($p in @($mergeStage,$existingZipDl)) {
                    if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
                    New-Item -ItemType Directory -Path $p -Force | Out-Null
                }
              $zipCounter = 0
              foreach ($z in $matchingZips) {
                    $zipCounter++
                    $localZip = Join-Path -Path $existingZipDl -ChildPath $z.Name
                    $zipExtract = Join-Path -Path $existingZipDl -ChildPath ("X{0:D2}" -f $zipCounter)

                    try {
                        Write-log ("Downloading ZIP for merge: {0}" -f $z.Name) -Sub -ToConsole -Severity INFO
                        Get-PnPFile -ServerRelativeUrl "$serverRelativeSubRoot/$($z.Name)" `
                                    -Path $existingZipDl `
                                    -FileName $($z.Name) `
                                    -AsFile -Force `
                                    -Connection $ActiveConnection

                        if (Test-Path -LiteralPath $zipExtract) { Remove-Item -LiteralPath $zipExtract -Recurse -Force -ErrorAction SilentlyContinue }
                        New-Item -ItemType Directory -Path $zipExtract -Force | Out-Null

                        Write-log ("Extracting {0} into {1}" -f $z.Name,$zipExtract) -Sub -ToConsole -Severity INFO
                        Expand-Archive -LiteralPath $localZip -DestinationPath $zipExtract -Force

                        # CHANGE alpha14: normalize each extracted ZIP before merge so folder structure stays inside one top-level container
                        if (-not $canonicalRoot) {
                            $canonicalZipName = Select-DuplicateZipOutputName -SourceFolderName $subitem.Name -MatchingZips $matchingZips -NonInteractive:$ProcessAll
                            $canonicalContainerName = [System.IO.Path]::GetFileNameWithoutExtension($canonicalZipName)
                            $canonicalRoot = Join-Path -Path $mergeStage -ChildPath $canonicalContainerName
                            if (-not (Test-Path -LiteralPath $canonicalRoot)) { New-Item -ItemType Directory -Path $canonicalRoot -Force | Out-Null }
                        }
                        $zipInfo = Get-NormalizedContainerInfo -Path $zipExtract -PreferredContainerName $canonicalContainerName
                        Merge-LocalFolderContents -SourceRoot $zipInfo.ContentRoot -TargetRoot $canonicalRoot -Mode $ConflictMode -LogPrefix ("[ZIP:{0}] " -f $z.Name) | Out-Null

                    } catch {
                        Write-log ("Failed to download/extract/merge {0}: {1}" -f $z.Name,$_.Exception.Message) -Sub -Severity ERROR -ToConsole
                        $ziperr++
                    } finally {
                        if (Test-Path -LiteralPath $zipExtract) { Remove-Item -LiteralPath $zipExtract -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                }

              # CHANGE alpha11/alpha12/alpha13/alpha14: let the user choose the final merged ZIP name when duplicate ZIPs already exist
              $canonicalZipName = Select-DuplicateZipOutputName -SourceFolderName $subitem.Name -MatchingZips $matchingZips -NonInteractive:$ProcessAll
              $canonicalContainerName = [System.IO.Path]::GetFileNameWithoutExtension($canonicalZipName)
              $canonicalRoot = Join-Path -Path $mergeStage -ChildPath $canonicalContainerName
              if (-not (Test-Path -LiteralPath $canonicalRoot)) { New-Item -ItemType Directory -Path $canonicalRoot -Force | Out-Null }

              try {
                    $sourceInfo = Get-NormalizedContainerInfo -Path $zipsourcefolder -PreferredContainerName $canonicalContainerName
                    Merge-LocalFolderContents -SourceRoot $sourceInfo.ContentRoot -TargetRoot $canonicalRoot -Mode $ConflictMode -LogPrefix "[NewDocs] " | Out-Null
                }
              catch {
                    Write-log ("Failed to merge new documents into staging: {0}" -f $_.Exception.Message) -Sub -Severity ERROR -ToConsole
                    $ziperr++
                }

              $zipdestination   = Join-Path -Path $workPaths.ZipRoot -ChildPath $canonicalZipName

              try {
                    if (Test-Path -LiteralPath $zipdestination) { Remove-Item -LiteralPath $zipdestination -Force -ErrorAction SilentlyContinue }
                    Write-log ("Creating merged ZIP {0} from staging {1}" -f $canonicalZipName,$mergeStage) -Sub -ToConsole -Severity INFO
                    Compress-Archive -Path (Join-Path $mergeStage '*') -DestinationPath $zipdestination -Force

                    $mergedCount = $null
                    $ZipFile = $null
                    try {
                        $ZipFile = [System.IO.Compression.ZipFile]::Open($zipdestination,[System.IO.Compression.ZipArchiveMode]::Read)
                        $mergedCount = $ZipFile.Entries.Count
                    } finally {
                        if ($ZipFile) { $ZipFile.Dispose() }
                    }
                    Write-log ("Merged ZIP created: {0} now holds {1} files" -f $canonicalZipName,$mergedCount) -Sub -ToConsole -Severity INFO

                }
              catch {
                    Write-log ("Failed to create merged ZIP: {0}" -f $_.Exception.Message) -Sub -Severity ERROR -ToConsole
                    $ziperr++
                }

              if ($ziperr -eq 0) {
                    $dupes = $matchingZips | Where-Object { $_.Name -ne $canonicalZipName }
                    foreach ($d in $dupes) {
                        Write-log ("Deleting duplicate ZIP on SharePoint (post-merge): {0}" -f $d.Name) -Sub -ToConsole -Severity WARN
                        if ($PSCmdlet.ShouldProcess("ZIP: $serverRelativeSubRoot/$($d.Name)","Remove-PnPFile (duplicate zip)")) {
                            try {
                                Remove-PnPFile -ServerRelativeUrl "$serverRelativeSubRoot/$($d.Name)" -Force -Connection $ActiveConnection
                            } catch {
                                Write-log ("Failed to delete duplicate ZIP {0}: {1}" -f $d.Name,$_.Exception.Message) -Sub -ToConsole -Severity WARN
                            }
                        }
                    }

                    $zipIndex[$refNumber] = @([pscustomobject]@{ Name = $canonicalZipName })
                }

              if (Test-Path -LiteralPath $mergeStage) { Remove-Item -LiteralPath $mergeStage -Recurse -Force -ErrorAction SilentlyContinue }
              if (Test-Path -LiteralPath $existingZipDl) { Remove-Item -LiteralPath $existingZipDl -Recurse -Force -ErrorAction SilentlyContinue }

            }
             else  # NO MATCHING ZIP, FIRST TIME COMPRESSION
             {
              Write-log "No matching ZIP found for reference $refNumber" -Sub -Severity INFO -ToConsole
              $zipdestination = Join-Path -Path $workPaths.ZipRoot -ChildPath "$($subitem.Name).zip"   # CHANGE alpha11/alpha12/alpha13/alpha14: create the new ZIP in the short ZIP work area
              Write-log "Creating $zipdestination" -Sub -Severity INFO -ToConsole
              try {
                    Compress-Archive -Path (Join-Path $zipsourcefolder '*') -DestinationPath $zipdestination -Force
                    $countNew = $null
                    $ZipFile = $null
                    try {
                        $ZipFile = [System.IO.Compression.ZipFile]::Open($zipdestination,[System.IO.Compression.ZipArchiveMode]::Read)
                        $countNew = $ZipFile.Entries.Count
                    } finally {
                        if ($ZipFile) { $ZipFile.Dispose() }
                    }
                    Write-log "Created new ZIP. File holds $countNew files" -Sub -Severity INFO -ToConsole
                }
              catch
               {
                Write-log "Failed to create new ZIP: $($_.Exception.Message)" -Sub -Severity ERROR -ToConsole
                $ziperr++
                }
               if ($ziperr -eq 0)
               {
                $zipIndex[$refNumber] = @([pscustomobject]@{ Name = (Split-Path -Leaf $zipdestination) })
               }

            }

             ###########################################################################################################
             #                                   UPLOAD FINAL ZIP BACK TO SHAREPOINT                                   #
             ###########################################################################################################

             $zipFileName = Split-Path -Leaf $zipdestination
             $zipServerRelativeUrl = "$serverRelativeSubRoot/$zipFileName"   # CHANGE alpha12/alpha13/alpha14: use the actual final ZIP name for checkout/checkin/upload
             #ZIP-centric upload log (ZIP is the entity; folder is the source input)
             Write-log ("[REF {0}] Uploading merged ZIP {1} to {2} (source folder: {3})" -f $refNumber,$zipFileName,$serverRelativeSubRoot,$subitem.Name) -Sub -Severity INFO -ToConsole

             # Alpha19 safety guard: re-check destination before upload so a stale inventory cannot overwrite an unrelated ZIP.
             $expectedExistingZipNames = @($matchingZips | ForEach-Object { $_.Name })
             if (-not (Test-DestinationZipUploadSafe -FolderSiteRelativeUrl $subrooturl `
                                                      -ServerRelativeSubRoot $serverRelativeSubRoot `
                                                      -ZipFileName $zipFileName `
                                                      -RefNumber $refNumber `
                                                      -ExpectedExistingZipNames $expectedExistingZipNames `
                                                      -Connection $ActiveConnection `
                                                      -Sub)) {
                 Write-log ("[REF {0}] Upload skipped; source folder kept untouched: {1}" -f $refNumber,$subitem.Name) -Sub -Severity WARN -ToConsole
                 continue
             }

             # If file already exists, try to checkout (ignore errors if not existing)
             try
               {
               Set-PnPFileCheckedOut -Url $zipServerRelativeUrl -Connection $ActiveConnection -ErrorAction SilentlyContinue
               } catch { }

             # Upload into site-relative folder (parent folder)
             $parentFolderSiteRel = $subrooturl.TrimStart("/")
             if (-not (Test-Path -LiteralPath $zipdestination)) { throw "Final ZIP for upload was not found locally: $zipdestination" }
             Add-PnPFile -Path $zipdestination -Folder $parentFolderSiteRel -Connection $ActiveConnection | Out-Null

             # Check-in (if it was checked out)
             try {
                 Set-PnPFileCheckedIn -Url $zipServerRelativeUrl -CheckinType MajorCheckIn -Connection $ActiveConnection -ErrorAction SilentlyContinue
                 } catch { }

             # Delete original folder (if allowed) with WhatIf support
             $FolderToDelete = Get-PnPFolder -Url $subitemUrl -Connection $ActiveConnection
             Write-log ("[REF {0}] Deleting source folder after successful archive: {1}" -f $refNumber,$FolderToDelete.Name) -Sub -Severity INFO -ToConsole  # 1805 CHANGE
             Remove-PnPFolderRecursive -Folder $FolderToDelete -ActiveConnection $ActiveConnection

             # finally delete the subitem itself
             $parentFolderForRemove = $subrooturl.TrimStart("/")
             if ($PSCmdlet.ShouldProcess("Folder: $($FolderToDelete.ServerRelativeUrl)","Remove-PnPFolder (root)"))
             {
             Remove-PnPFolder -Name $subitem.Name -Folder $parentFolderForRemove -Force -Connection $ActiveConnection
             }
              
         } # END FOR EACH SUBITEM (FOLDER) IN THE CURRENT SELECTION OF FOLDERS (example: Folder 964000 - 964999)

     # Upload per-root log after this range folder has been processed.
     # CHANGE alpha16: this remains inside the root-folder loop so every selected range gets its own uploaded log.
     if (Test-Path -LiteralPath $subLogPath)
      {
       try
        {
         $logDestSiteRel = "$subrooturl/$LogFolderName"
         $logDestFolder = Get-PnPFolder -Url $logDestSiteRel -Connection $ActiveConnection
         Write-log "Uploading compression log $sublogfile to $($logDestFolder.ServerRelativeUrl)" -Severity INFO -ToConsole  # 1805 CHANGE
         Add-PnPFile -Path $subLogPath -Folder $logDestSiteRel.TrimStart("/") -Connection $ActiveConnection | Out-Null
        }
       catch
        {
         Write-log "Failed to upload compression log $sublogfile : $($_.Exception.Message)" -Severity WARN -ToConsole
        }
      }

    } # END foreach ($rootfold in $selected)

} # END PROCESS

end

{
# Clean this run's temporary folder 
if (Test-Path -LiteralPath $RunTemp)
{
  try
  {
   Get-ChildItem $RunTemp -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
   Remove-Item $RunTemp -Force -ErrorAction SilentlyContinue
  } catch { }
}
    Write-Host "Execution complete"
}
