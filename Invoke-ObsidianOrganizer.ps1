<#
.SYNOPSIS
    Organizes an Obsidian vault by sending new/un-processed notes to Claude
    Haiku 4.5 (via the Claude Code CLI) to be rewritten, tagged, flagged, and
    cross-linked. Built to run headlessly from Windows Task Scheduler.

.DESCRIPTION
    For each eligible markdown note the script:
      1. Backs up the original to <name>.original.md (in the vault)
      2. Sends the content to Claude Haiku 4.5 with instructions to rewrite,
         tag by topic, add #priority-brainstorm to brainstorming notes, and
         append a "## Related Topics" section
      3. Writes the cleaned note back to the vault (UTF-8, no BOM)
      4. Records the note in processed.json so it is never reprocessed

    Uses the Claude Code CLI (`claude -p`) for inference — no API key needed.
    Auth is read from stored Claude Code credentials (~/.claude/).

    Safety features:
      * Per-note .original.md backup before any change
      * Skip-guard: notes already containing "## Related Topics" are skipped
      * processed.json tracks handled notes permanently
      * -WhatIf previews actions without changing anything
      * Failures are logged and the note stays in the queue for next run
      * .original.md backups and excluded folders are never processed

.PARAMETER VaultPath
    Full path to the Obsidian vault folder (inside iCloud Drive).

.PARAMETER Model
    Anthropic model ID. Defaults to claude-haiku-4-5-20251001.

.PARAMETER ExcludeFolders
    Top-level vault subfolders to skip (auto-generated logs by default).

.PARAMETER MaxNotes
    Maximum notes to process in a single run (throttle / cost guard). 0 = no limit.

.PARAMETER MaxNoteChars
    Maximum note length in characters. Notes longer than this are skipped with a
    WARN log and NOT recorded in processed.json, so they are retried automatically
    if the threshold is later raised. Default 50000. Set to 0 to disable the guard.

.PARAMETER WhatIf
    Preview mode. Logs what would happen but writes nothing.

.PARAMETER KeepBackups
    When specified, the per-note .original.md backup created before each rewrite
    is retained in the vault. By default (without this switch) the backup is
    deleted immediately after a successful rewrite.

.EXAMPLE
    .\Invoke-ObsidianOrganizer.ps1 -VaultPath "C:\Users\Brian\iCloudDrive\iCloud~md~obsidian\Brainstorming"

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $VaultPath,

    [string] $Model = "claude-haiku-4-5-20251001",

    [string[]] $ExcludeFolders = @("sessions", "conversations", "NEXUS"),

    [int] $MaxNotes = 0,

    [int] $MaxNoteChars = 50000,

    [switch] $WhatIf,

    [switch] $Force,

    [switch] $FileIntoFolders,

    [switch] $RefileOnly,

    [switch] $KeepBackups
)

# ----------------------------------------------------------------------------
# Setup
# ----------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile     = Join-Path $ScriptDir "organizer.log"
$ProcessedDb = Join-Path $ScriptDir "processed.json"
$Categories  = @("Automation", "Homelab", "Projects", "Reference", "Tools", "Work")
$ClaudeExe   = if (Test-Path "C:\Users\Brian\AppData\Roaming\npm\claude.cmd") { "C:\Users\Brian\AppData\Roaming\npm\claude.cmd" } else { "claude" }

# LLM semantic category -> PARA lifecycle folder
$ParaFolderMap = @{
    "Automation" = "20-Areas"
    "Homelab"    = "20-Areas"
    "Projects"   = "10-Projects"
    "Reference"  = "30-Resources"
    "Tools"      = "30-Resources"
    "Work"       = "20-Areas"
}
# LLM semantic category -> PARA status frontmatter value
$ParaStatusMap = @{
    "Automation" = "active"
    "Homelab"    = "active"
    "Projects"   = "active"
    "Reference"  = "evergreen"
    "Tools"      = "evergreen"
    "Work"       = "active"
}
# All PARA top-level folders this organizer manages (never re-file notes already here)
$ParaFolders = @("00-Inbox", "10-Projects", "20-Areas", "30-Resources", "40-Archive", "MOCs")
$Utf8NoBom   = New-Object System.Text.UTF8Encoding($false)

# Log rotation: keep last 2000 lines
if (Test-Path $LogFile) {
    $logLines = [System.IO.File]::ReadAllLines($LogFile)
    if ($logLines.Count -gt 2000) {
        $trimmed = $logLines[($logLines.Count - 2000)..($logLines.Count - 1)]
        [System.IO.File]::WriteAllLines($LogFile, $trimmed, $Utf8NoBom)
    }
}

function Write-Log {
    param([string] $Message, [string] $Level = "INFO")
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line  = "{0} [{1}] {2}" -f $stamp, $Level, $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Host $line
}

function Read-TextFile {
    param([string] $Path)
    return [System.IO.File]::ReadAllText($Path)
}

function Write-TextFile {
    param([string] $Path, [string] $Content)
    [System.IO.File]::WriteAllText($Path, $Content, $Utf8NoBom)
}

function Show-ToastNotification {
    param(
        [string] $Title,
        [string] $Message
    )
    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.UI.Notifications.ToastNotification,        Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument,                  Windows.Data.Xml.Dom,     ContentType = WindowsRuntime]

        $appId    = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
        $escTitle = [System.Security.SecurityElement]::Escape($Title)
        $escMsg   = [System.Security.SecurityElement]::Escape($Message)

        $xml = @"
<toast>
  <visual>
    <binding template="ToastGeneric">
      <text>$escTitle</text>
      <text>$escMsg</text>
    </binding>
  </visual>
</toast>
"@
        $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $doc.LoadXml($xml)
        $toast = New-Object Windows.UI.Notifications.ToastNotification($doc)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
    }
    catch {
        Write-Log "Toast notification failed (non-fatal): $($_.Exception.Message)" "WARN"
    }
}

# ----------------------------------------------------------------------------
# Verify Claude CLI is available
# ----------------------------------------------------------------------------
if (-not $WhatIf) {
    $claudeFound = (Get-Command $ClaudeExe -ErrorAction SilentlyContinue) -or (Test-Path $ClaudeExe -ErrorAction SilentlyContinue)
    if (-not $claudeFound) {
        Write-Log "Claude CLI not found at: $ClaudeExe — install Claude Code or fix the path." "ERROR"
        exit 1
    }
}

# Normalize VaultPath: strip trailing separator to prevent off-by-one in relative path calculations
$VaultPath = $VaultPath.TrimEnd('\', '/')

Write-Log "===== Obsidian Organizer run started (WhatIf=$($WhatIf.IsPresent)) ====="

if (-not (Test-Path -LiteralPath $VaultPath)) {
    Write-Log "Vault path not found: $VaultPath" "ERROR"
    exit 1
}

# ----------------------------------------------------------------------------
# Load processed database
# ----------------------------------------------------------------------------
$processed = @{}
if (Test-Path $ProcessedDb) {
    try {
        $raw = Read-TextFile $ProcessedDb | ConvertFrom-Json
        foreach ($prop in $raw.PSObject.Properties) {
            $processed[$prop.Name] = $prop.Value
        }
    } catch {
        Write-Log "processed.json unreadable, starting fresh: $($_.Exception.Message)" "WARN"
        $processed = @{}
    }
}

function Save-Processed {
    if ($WhatIf) { return }
    ($processed | ConvertTo-Json -Depth 5) | ForEach-Object { Write-TextFile $ProcessedDb $_ }
}

# ----------------------------------------------------------------------------
# Gather candidate notes
# ----------------------------------------------------------------------------
$excludeSet = @{}
foreach ($f in $ExcludeFolders) { $excludeSet[$f.ToLowerInvariant()] = $true }

$allMd = Get-ChildItem -LiteralPath $VaultPath -Filter *.md -Recurse -File

$candidates = foreach ($file in $allMd) {
    # Skip our own backups
    if ($file.Name -like "*.original.md") { continue }

    # Skip date-named files (daily briefings, session files: YYYY-MM-DD.md or YYYY-MM-DD-*.md)
    if ($file.Name -match '^\d{4}-\d{2}-\d{2}[-.]') { continue }

    # Relative path from vault root
    $rel = $file.FullName.Substring($VaultPath.Length).TrimStart('\','/')

    # Skip excluded top-level folders
    $topFolder = ($rel -split '[\\/]')[0]
    if ($rel -match '[\\/]' -and $excludeSet.ContainsKey($topFolder.ToLowerInvariant())) { continue }

    # Skip already processed (unless -Force)
    if (-not $Force -and $processed.ContainsKey($rel)) { continue }

    $file
}

$candidates = @($candidates)
Write-Log ("Found {0} candidate note(s) to evaluate." -f $candidates.Count)

# ----------------------------------------------------------------------------
# Build a list of existing note titles (for Related Topics cross-linking)
# ----------------------------------------------------------------------------
$noteTitles = ($allMd |
    Where-Object { $_.Name -notlike "*.original.md" } |
    ForEach-Object { $_.BaseName } |
    Sort-Object -Unique) -join ", "

# ----------------------------------------------------------------------------
# Prompt template
# ----------------------------------------------------------------------------
$systemPrompt = @"
You are an expert note organizer for an Obsidian vault. You receive the raw
content of a single note and return a clean, rewritten version.

RULES:
1. Begin the output with a YAML frontmatter block delimited by lines of exactly
   three dashes (---). It must contain a single line:
       category: <Category>
   You MUST choose exactly one of these six categories, copied verbatim:
     - Automation  : scripts, Task Scheduler jobs, CI/CD, scheduled automation, PowerShell tools
     - Homelab     : homelab servers and infrastructure (Hermes, Unraid, NEXUS, networking, Home Assistant)
     - Projects    : active in-progress project work and project session notes
     - Reference   : SOPs, setup/install guides, cheat sheets, architecture docs (stable how-to material)
     - Tools       : CLI tools and external tool documentation (codex-verify, codex CLI, openrouter, etc.)
     - Work        : employer/job-related content (GM, Ford, BeyondTrust)
   Apply these rules in order -- first match wins:
     a. Homelab infrastructure, servers, or networking (Hermes, Unraid, NEXUS, Home Assistant) -> Homelab
     b. CLI or external tool documentation -> Tools
     c. Automation script, Task Scheduler, or CI/CD note -> Automation
     d. Employer-related content (GM, Ford, BeyondTrust) -> Work
     e. Stable how-to, SOP, setup guide, or reference cheat sheet -> Reference
     f. Active in-progress project work -> Projects
   If genuinely uncertain after applying these rules, write: category: UNSORTED
   Do NOT add a date, created, updated, or any other field to the frontmatter.
   Output only the category line (or UNSORTED) inside the frontmatter.
2. Immediately AFTER the closing --- of the frontmatter, put a single line of
   inline Obsidian tags using the #tag-name form. Include:
     - the main topics/subjects of the note
     - an #app-name style tag when a specific app or project is discussed
     - a #category/<slug> tag matching the category from rule 1, where <slug>
       is the category lowercased with spaces replaced by hyphens
       (e.g. category "Homelab" -> #category/homelab)
     - #priority-brainstorm if the note is a brainstorming session
3. Rewrite the body into clear, well-structured Markdown. Fix grammar and flow,
   use headings, lists, and short paragraphs. NEVER change the meaning, drop
   information, or invent facts. Preserve every concrete detail.
4. End the note with a section exactly titled "## Related Topics" that lists
   subjects or other notes this note likely connects to. When relevant, link to
   existing notes using [[Note Title]] wiki-link syntax. Existing note titles:
   {NOTE_TITLES}
5. Output ONLY the final Markdown for the note (starting with the frontmatter).
   No preamble, no explanation, and do NOT wrap the whole thing in a code fence.
"@
$systemPrompt = $systemPrompt.Replace("{NOTE_TITLES}", $noteTitles)

function Add-FrontmatterFields {
    param(
        [string]   $Markdown,
        [datetime] $Date,
        [string]   $Status  = "",
        [string]   $Updated = ""
    )

    $dateStr    = $Date.ToString("yyyy-MM-dd")
    $updatedStr = if ($Updated) { $Updated } else { (Get-Date).ToString("yyyy-MM-dd") }
    $pattern    = '(?s)\A(---\r?\n)(.*?)(\r?\n---\s*?(\r?\n|$))'

    function Set-Field([string]$body, [string]$key, [string]$val) {
        if ($body -match "(?m)^\s*${key}:\s*.*$") {
            return [System.Text.RegularExpressions.Regex]::Replace($body, "(?m)^\s*${key}:\s*.*$", "${key}: $val")
        }
        return $body.TrimEnd() + "`n${key}: $val"
    }

    if ($Markdown -match $pattern) {
        $open = $matches[1]
        $body = $matches[2]
        $body = Set-Field $body "date"    $dateStr
        $body = Set-Field $body "updated" $updatedStr
        if (-not [string]::IsNullOrWhiteSpace($Status)) {
            $body = Set-Field $body "status" $Status
        }
        $close = $matches[3]
        $rest  = $Markdown.Substring($matches[0].Length)
        return $open + $body + $close + $rest
    }

    # No frontmatter found: prepend a minimal block
    $fm = "---`ndate: $dateStr`nupdated: $updatedStr"
    if (-not [string]::IsNullOrWhiteSpace($Status)) { $fm += "`nstatus: $Status" }
    return $fm + "`n---`n`n" + $Markdown
}

function Invoke-ClaudeCode {
    param([string] $NoteContent)

    $userMessage = "Here is the note to rewrite:`n`n$NoteContent"
    $retryDelays = @(5, 15)
    $attempt = 0
    while ($true) {
        try {
            $responseJson = $userMessage | & $ClaudeExe -p --system-prompt-file $SystemPromptFile --model $Model --output-format json 2>&1
            $response = $responseJson | ConvertFrom-Json
            if ($response.is_error) { throw "Claude CLI error: $($response.result)" }
            return @{
                Content = $response.result
                CostUsd = [double]($response.total_cost_usd)
            }
        } catch {
            if ($attempt -lt $retryDelays.Count) {
                $delay = $retryDelays[$attempt]
                Write-Log "Claude CLI error (attempt $($attempt + 1)): $($_.Exception.Message). Retrying in ${delay}s..." "WARN"
                Start-Sleep -Seconds $delay
                $attempt++
            } else {
                throw
            }
        }
    }
}

# ----------------------------------------------------------------------------
# -RefileOnly mode: move notes to PARA folders using existing category: frontmatter
# No API calls — reads what Claude already assigned and maps to PARA names.
# Run once after the PARA folder names are adopted to migrate existing notes.
# ----------------------------------------------------------------------------
if ($RefileOnly) {
    Write-Log "===== RefileOnly mode: moving notes to PARA folders based on existing category: frontmatter ====="
    $refileMoved   = 0
    $refileSkipped = 0

    foreach ($file in $allMd) {
        if ($file.Name -like "*.original.md") { continue }
        $rel       = $file.FullName.Substring($VaultPath.Length).TrimStart('\','/')
        $topFolder = ($rel -split '[\\/]')[0]
        if ($rel -match '[\\/]' -and $excludeSet.ContainsKey($topFolder.ToLowerInvariant())) { continue }

        try {
            $noteContent = Read-TextFile $file.FullName
            $cat = $null
            if ($noteContent -match '(?ms)^---\s*\r?\n.*?^category:\s*(.+?)\s*\r?\n.*?^---\s*$') {
                $cat = $matches[1].Trim().Trim('"').Trim("'")
            }
            if ([string]::IsNullOrWhiteSpace($cat) -or $cat -eq "UNSORTED" -or $Categories -notcontains $cat) {
                $refileSkipped++
                continue
            }

            $paraFolder = $ParaFolderMap[$cat]
            $targetDir  = Join-Path $VaultPath $paraFolder

            if ($file.DirectoryName.TrimEnd('\') -eq $targetDir.TrimEnd('\')) {
                Write-Log "REFILE: already in $paraFolder, skipping: $rel"
                $refileSkipped++
                continue
            }

            if ($WhatIf) {
                Write-Log "REFILE WHATIF: $rel -> $paraFolder\"
                $refileMoved++
                continue
            }

            if (-not (Test-Path -LiteralPath $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            $destNote = Join-Path $targetDir $file.Name
            if (Test-Path -LiteralPath $destNote) {
                $sfx = 2
                do {
                    $destNote = Join-Path $targetDir ($file.BaseName + "_$sfx" + $file.Extension)
                    $sfx++
                } while ((Test-Path -LiteralPath $destNote) -and $sfx -le 99)
                Write-Log "REFILE: name collision resolved, using: $(Split-Path $destNote -Leaf)" "WARN"
            }
            Move-Item -LiteralPath $file.FullName -Destination $destNote
            if (Test-Path -LiteralPath $file.FullName) {
                Remove-Item -LiteralPath $file.FullName -Force
                Write-Log "REFILE CLEANUP: removed leftover source: $rel" "WARN"
            }
            # Update processed.json to the new relative path
            $newRel = $destNote.Substring($VaultPath.Length).TrimStart('\','/')
            if ($processed.ContainsKey($rel)) {
                $processed.Remove($rel) | Out-Null
                $processed[$newRel] = @{ status = "organized"; at = (Get-Date).ToString("o"); category = $cat }
            }
            Write-Log "REFILE: $rel -> $newRel"
            $refileMoved++
        } catch {
            Write-Log "REFILE ERROR: $rel : $($_.Exception.Message)" "ERROR"
        }
    }

    Save-Processed
    Write-Log ("RefileOnly complete. moved={0} skipped={1}" -f $refileMoved, $refileSkipped)
    Write-Log "===== Obsidian Organizer run finished (RefileOnly) ====="
    if (-not $WhatIf) {
        Show-ToastNotification -Title "Obsidian Refile complete" -Message ("Moved: {0}  Skipped: {1}" -f $refileMoved, $refileSkipped)
    }
    exit 0
}

# ----------------------------------------------------------------------------
# Write system prompt to temp file (reused across all notes in this run)
# ----------------------------------------------------------------------------
$SystemPromptFile = [System.IO.Path]::GetTempFileName() + ".md"
[System.IO.File]::WriteAllText($SystemPromptFile, $systemPrompt, $Utf8NoBom)

# ----------------------------------------------------------------------------
# Process notes
# ----------------------------------------------------------------------------
$count = 0
$succeeded = 0
$skipped = 0
$failed = 0
$totalCostUsd = 0.0

foreach ($file in $candidates) {
    if ($MaxNotes -gt 0 -and $count -ge $MaxNotes) {
        Write-Log "Reached MaxNotes limit ($MaxNotes). Remaining notes will run next time."
        break
    }
    $count++

    $rel = $file.FullName.Substring($VaultPath.Length).TrimStart('\','/')

    try {
        $content = Read-TextFile $file.FullName

        # Skip-guard: already organized (bypassed by -Force)
        if (-not $Force -and $content -match '(?m)^\s*##\s+Related Topics\s*$') {
            Write-Log "SKIP (already has Related Topics): $rel"
            $processed[$rel] = @{ status = "skipped"; at = (Get-Date).ToString("o") }
            $skipped++
            continue
        }

        if ([string]::IsNullOrWhiteSpace($content)) {
            Write-Log "SKIP (empty file): $rel"
            $processed[$rel] = @{ status = "skipped-empty"; at = (Get-Date).ToString("o") }
            $skipped++
            continue
        }

        # Size guard: skip oversized notes WITHOUT recording them so they retry if -MaxNoteChars is raised
        if ($MaxNoteChars -gt 0 -and $content.Length -gt $MaxNoteChars) {
            Write-Log ("SKIP (too large: {0} chars > MaxNoteChars {1}): {2}" -f $content.Length, $MaxNoteChars, $rel) "WARN"
            $skipped++
            continue
        }

        if ($WhatIf) {
            Write-Log "WHATIF would process: $rel ($($content.Length) chars)"
            $succeeded++
            continue
        }

        Write-Log "Processing: $rel"
        $result    = Invoke-ClaudeCode -NoteContent $content
        $rewritten = $result.Content
        $totalCostUsd += $result.CostUsd
        Write-Log ("COST: `${0:F6} for: {1}" -f $result.CostUsd, $rel)

        if ([string]::IsNullOrWhiteSpace($rewritten)) {
            throw "Empty rewrite returned."
        }

        # Determine PARA status from the category Claude assigned
        $earlyStatus = ""
        if ($rewritten -match '(?ms)^---\s*\r?\n.*?^category:\s*(.+?)\s*\r?\n.*?^---\s*$') {
            $earlyCat = $matches[1].Trim().Trim('"').Trim("'")
            if ($ParaStatusMap.ContainsKey($earlyCat)) { $earlyStatus = $ParaStatusMap[$earlyCat] }
        }

        # Inject date (original), updated (today), and status into frontmatter.
        # Capture LastWriteTime before overwriting the file.
        $originalDate = $file.LastWriteTime
        $todayStr     = (Get-Date).ToString("yyyy-MM-dd")
        $rewritten    = Add-FrontmatterFields -Markdown $rewritten -Date $originalDate -Status $earlyStatus -Updated $todayStr
        Write-Log ("DATE: injected date={0} updated={1} status={2} into: {3}" -f $originalDate.ToString("yyyy-MM-dd"), $todayStr, $earlyStatus, $rel)

        # Backup original (only if a backup does not already exist)
        $backupPath = Join-Path $file.DirectoryName ($file.BaseName + ".original.md")
        if (-not (Test-Path $backupPath)) {
            Write-TextFile $backupPath $content
        }

        # Write cleaned note
        Write-TextFile $file.FullName $rewritten

        $processed[$rel] = @{ status = "organized"; at = (Get-Date).ToString("o") }
        Save-Processed
        Write-Log "DONE: $rel"
        $succeeded++

        # Remove the backup unless -KeepBackups was specified
        if (-not $KeepBackups -and (Test-Path -LiteralPath $backupPath)) {
            Remove-Item -LiteralPath $backupPath -Force
            Write-Log "BACKUP DELETED: $($file.BaseName).original.md"
        }

        # Optionally file the note into its PARA folder (mapped from semantic category)
        if ($FileIntoFolders) {
            $cat = $null
            if ($rewritten -match '(?ms)^---\s*\r?\n.*?^category:\s*(.+?)\s*\r?\n.*?^---\s*$') {
                $cat = $matches[1].Trim().Trim('"').Trim("'")
            }
            if ([string]::IsNullOrWhiteSpace($cat)) {
                Write-Log "FILE: no category found, leaving in place: $rel" "WARN"
            } elseif ($cat -eq "UNSORTED") {
                Write-Log "FILE: model returned UNSORTED, leaving in place for manual review: $rel" "WARN"
            } elseif ($Categories -notcontains $cat) {
                Write-Log "FILE: model returned unknown category '$cat' (not in allowlist), leaving in place: $rel" "WARN"
            } else {
                $paraFolder       = $ParaFolderMap[$cat]
                $targetDir        = Join-Path $VaultPath $paraFolder
                $currentTopFolder = ($rel -split '[\\/]')[0]
                $alreadyCorrect   = ($file.DirectoryName.TrimEnd('\') -eq $targetDir.TrimEnd('\'))

                if ($alreadyCorrect) {
                    Write-Log "FILE: already in correct PARA folder ($paraFolder), no move needed: $rel"
                } elseif ($ParaFolders -contains $currentTopFolder) {
                    Write-Log "FILE: note is in PARA folder '$currentTopFolder' but maps to '$paraFolder' -- skipping to avoid misfile: $rel" "WARN"
                } else {
                    if (-not (Test-Path -LiteralPath $targetDir)) {
                        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                    }
                    $destNote = Join-Path $targetDir $file.Name
                    # Resolve name collision by appending _2, _3, ... instead of skipping
                    if (Test-Path -LiteralPath $destNote) {
                        $sfx = 2
                        do {
                            $destNote = Join-Path $targetDir ($file.BaseName + "_$sfx" + $file.Extension)
                            $sfx++
                        } while ((Test-Path -LiteralPath $destNote) -and $sfx -le 99)
                        Write-Log "FILE: name collision resolved, using: $(Split-Path $destNote -Leaf)" "WARN"
                    }
                    Move-Item -LiteralPath $file.FullName -Destination $destNote
                    # Defensive guard: ensure no source copy remains after the move
                    if (Test-Path -LiteralPath $file.FullName) {
                        Remove-Item -LiteralPath $file.FullName -Force
                        Write-Log "CLEANUP: removed leftover source after filing: $rel" "WARN"
                    }
                    # Move the backup alongside if it still exists
                    if (Test-Path -LiteralPath $backupPath) {
                        $destBak = Join-Path $targetDir (Split-Path $backupPath -Leaf)
                        if (-not (Test-Path -LiteralPath $destBak)) {
                            Move-Item -LiteralPath $backupPath -Destination $destBak
                        }
                    }
                    # Update processed.json to the note's new relative path
                    $newRel = $destNote.Substring($VaultPath.Length).TrimStart('\','/')
                    $processed.Remove($rel) | Out-Null
                    $processed[$newRel] = @{ status = "organized"; at = (Get-Date).ToString("o"); category = $cat }
                    Save-Processed
                    Write-Log "FILED: $rel -> $newRel"
                }
            }
        }
    }
    catch {
        Write-Log "ERROR processing $rel : $($_.Exception.Message)" "ERROR"
        $failed++
        # Note intentionally left out of processed{} so it retries next run
    }
}

Save-Processed
Write-Log ("Run complete. processed={0} succeeded={1} skipped={2} failed={3} cost=`${4:F4}" -f $count, $succeeded, $skipped, $failed, $totalCostUsd)
Write-Log "===== Obsidian Organizer run finished ====="

Remove-Item $SystemPromptFile -Force -ErrorAction SilentlyContinue

if (-not $WhatIf) {
    $toastTitle = "Obsidian Organizer finished"
    $toastBody  = ("Succeeded: {0}  Skipped: {1}  Failed: {2}`nTotal cost: `${3:F4}" -f $succeeded, $skipped, $failed, $totalCostUsd)
    Show-ToastNotification -Title $toastTitle -Message $toastBody
}
