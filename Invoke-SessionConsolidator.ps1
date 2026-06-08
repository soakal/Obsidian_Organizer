<#
.SYNOPSIS
    Daily session consolidator for the Obsidian vault.

.DESCRIPTION
    Runs at 11:59 PM daily.
    1. Finds all session files from today in sessions/
    2. Merges them into a single dated note (sessions/YYYY-MM-DD.md)
    3. Deletes the individual session files
    4. Uses Claude to identify which existing vault documents should be updated
    5. Rewrites those documents with the new information

.PARAMETER VaultPath
    Full path to the Obsidian vault folder.

.PARAMETER ApiKey
    OpenRouter API key. Reads from OPENROUTER_API_KEY env var or apikey.txt if omitted.

.PARAMETER Model
    OpenRouter model slug. Defaults to anthropic/claude-haiku-4.5.

.PARAMETER MaxDocUpdates
    Maximum vault documents to update per run. Default: 5.

.PARAMETER WhatIf
    Preview only -- no files are written or deleted.
#>

[CmdletBinding()]
param(
    [string] $VaultPath = "C:\Users\Brian\iCloudDrive\iCloud~md~obsidian\Brainstorming",
    [string] $ApiKey,
    [string] $Model = "anthropic/claude-haiku-4.5",
    [int]    $MaxDocUpdates = 5,
    [int]    $LookbackDays  = 3,
    [switch] $WhatIf
)

$ErrorActionPreference = "Stop"
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile    = Join-Path $ScriptDir "consolidator.log"
$Endpoint   = "https://openrouter.ai/api/v1/chat/completions"
$Utf8NoBom  = New-Object System.Text.UTF8Encoding($false)
$Today      = (Get-Date).ToString("yyyy-MM-dd")

# C1: Force TLS 1.2 -- PS 5.1 defaults to TLS 1.0 which OpenRouter rejects
[Net.ServicePointManager]::SecurityProtocol = `
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# --- Logging ----------------------------------------------------------------

function Write-Log {
    param([string] $Message, [string] $Level = "INFO")
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line  = "{0} [{1}] {2}" -f $stamp, $Level, $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Host $line
}

# Trim log to last 2000 lines
if (Test-Path $LogFile) {
    $lines = Get-Content $LogFile -Encoding UTF8
    if ($lines.Count -gt 2000) {
        $lines | Select-Object -Last 2000 | Set-Content $LogFile -Encoding UTF8
    }
}

# --- File helpers -----------------------------------------------------------

function Read-TextFile {
    param([string] $Path)
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Write-TextFile {
    param([string] $Path, [string] $Content)
    [System.IO.File]::WriteAllText($Path, $Content, $Utf8NoBom)
}

function Get-ConsolidatedSessions {
    param([string] $FilePath)
    if (-not (Test-Path -LiteralPath $FilePath)) { return @() }
    $text = Read-TextFile $FilePath
    if ($text -match '(?m)^consolidated_sessions:\s*(\[.*?\])') {
        try { return @($matches[1] | ConvertFrom-Json) } catch {}
    }
    return @()
}

function Set-ConsolidatedSessions {
    param([string] $FilePath, [string[]] $Sessions)
    $text = Read-TextFile $FilePath
    $json = "[" + (($Sessions | ForEach-Object { '"' + $_ + '"' }) -join ",") + "]"
    if ($text -match '(?s)\A(---\r?\n)(.*?)(\r?\n---\r?\n)') {
        $fm = $matches[2]
        if ($fm -match '(?m)^consolidated_sessions:') {
            $fm = ($fm -split "`r?`n" | ForEach-Object {
                if ($_ -match '^consolidated_sessions:') { "consolidated_sessions: $json" } else { $_ }
            }) -join "`n"
        } else {
            $fm = $fm.TrimEnd() + "`nconsolidated_sessions: $json"
        }
        $text = "---`n" + $fm + "`n---`n" + $text.Substring($matches[0].Length)
    }
    Write-TextFile $FilePath $text
}

# --- API call ---------------------------------------------------------------

function Invoke-Claude {
    param(
        [string] $Prompt,
        [string] $SystemMsg = "You are a helpful assistant."
    )
    $body = @{
        model       = $Model
        messages    = @(
            @{ role = "system"; content = $SystemMsg },
            @{ role = "user";   content = $Prompt }
        )
        temperature = 0.2
    } | ConvertTo-Json -Depth 8

    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type"  = "application/json"
        "HTTP-Referer"  = "https://github.com/local/obsidian-consolidator"
        "X-Title"       = "Obsidian Session Consolidator"
    }

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $resp  = Invoke-RestMethod -Uri $Endpoint -Method Post -Headers $headers `
                               -Body $bytes -ContentType "application/json; charset=utf-8"

    if (-not $resp.choices -or $resp.choices.Count -eq 0) {
        throw "API returned no choices."
    }
    return $resp.choices[0].message.content
}

# --- Resolve API key --------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    if (-not [string]::IsNullOrWhiteSpace($env:OPENROUTER_API_KEY)) {
        $ApiKey = $env:OPENROUTER_API_KEY
    } else {
        $keyFile = Join-Path $ScriptDir "apikey.txt"
        if (Test-Path $keyFile) {
            $ApiKey = (Get-Content $keyFile |
                Where-Object { $_.Trim() -ne "" -and -not $_.TrimStart().StartsWith("#") } |
                Select-Object -First 1)
            if ($ApiKey) { $ApiKey = $ApiKey.Trim() }
        }
    }
}

Write-Log "===== Session Consolidator started (WhatIf=$($WhatIf.IsPresent), Date=$Today) ====="

if ([string]::IsNullOrWhiteSpace($ApiKey) -and -not $WhatIf) {
    Write-Log "No API key found. Set OPENROUTER_API_KEY or create apikey.txt." "ERROR"
    exit 1
}

if (-not (Test-Path -LiteralPath $VaultPath)) {
    Write-Log "Vault path not found: $VaultPath" "ERROR"
    exit 1
}

# =============================================================================
# STEPS 1-3 -- Find, merge, and write session files (N-day lookback)
# =============================================================================

$SessionsDir = Join-Path $VaultPath "sessions"
if (-not (Test-Path -LiteralPath $SessionsDir)) {
    Write-Log "Sessions folder not found: $SessionsDir" "WARN"
    Write-Log "===== Session Consolidator finished ====="
    exit 0
}

$allSessionBodies   = [System.Collections.Generic.List[string]]::new()
$consolidatedDates  = [System.Collections.Generic.List[string]]::new()
$totalSessionsFound = 0

for ($daysBack = 0; $daysBack -le $LookbackDays; $daysBack++) {
    $targetDate = (Get-Date).AddDays(-$daysBack).ToString("yyyy-MM-dd")

    # Per-session files: YYYY-MM-DD-SomeName.md (date-prefix + hyphen + title)
    # Bare YYYY-MM-DD.md consolidated daily files are excluded by requiring -.+
    $allDateFiles = @(Get-ChildItem -LiteralPath $SessionsDir -Filter "*.md" -File |
        Where-Object { $_.Name -match "^$targetDate-.+\.md$" } |
        Sort-Object Name)

    if ($allDateFiles.Count -eq 0) { continue }

    $consolidatedPath = Join-Path $SessionsDir "$targetDate.md"
    $consolidatedRel  = "sessions\$targetDate.md"

    # Idempotency: filter out sessions already recorded in the daily file's frontmatter
    $alreadyMerged = @(Get-ConsolidatedSessions -FilePath $consolidatedPath)
    $dateFiles     = @($allDateFiles | Where-Object { $alreadyMerged -notcontains $_.Name })

    if ($dateFiles.Count -eq 0) {
        Write-Log "All $($allDateFiles.Count) session(s) for $targetDate already consolidated -- skipping."
        continue
    }

    Write-Log "Found $($dateFiles.Count) new session file(s) for $targetDate ($($alreadyMerged.Count) already merged)."
    $totalSessionsFound += $dateFiles.Count

    $sessionParts = foreach ($file in $dateFiles) {
        $raw   = Read-TextFile $file.FullName
        $body  = $raw -replace '(?s)\A---\r?\n.*?\r?\n---\r?\n', ''
        $title = $file.BaseName -replace "^$targetDate-", '' -replace '-', ' '
        "### $title`n`n$($body.Trim())"
    }
    $dateCombinedBody = $sessionParts -join "`n`n---`n`n"
    $allSessionBodies.Add($dateCombinedBody)

    # Full merged list for writing to frontmatter
    $mergedNames = $alreadyMerged + @($dateFiles | ForEach-Object { $_.Name })

    if (Test-Path -LiteralPath $consolidatedPath) {
        if ($WhatIf) {
            Write-Log "WHATIF: Would append $($dateFiles.Count) session(s) to existing $consolidatedRel"
        } else {
            $existing = Read-TextFile $consolidatedPath
            $appended = $existing.TrimEnd() + "`n`n---`n`n" + $dateCombinedBody
            Write-TextFile $consolidatedPath $appended
            Set-ConsolidatedSessions -FilePath $consolidatedPath -Sessions $mergedNames
            Write-Log "Appended $($dateFiles.Count) session(s) to existing: $consolidatedRel"
        }
    } else {
        $sessionsJson = "[" + (($mergedNames | ForEach-Object { '"' + $_ + '"' }) -join ",") + "]"
        $newContent = @"
---
date: "$targetDate"
type: "daily-session"
consolidated_sessions: $sessionsJson
tags: ["session", "daily-log", "category/sessions"]
---

#session #daily-log #category/sessions

# Session Log -- $targetDate

$dateCombinedBody

## Related Topics

- [[All-Sessions]]
"@
        if ($WhatIf) {
            Write-Log "WHATIF: Would write $consolidatedRel"
        } else {
            Write-TextFile $consolidatedPath $newContent
            Write-Log "Wrote: $consolidatedRel"
        }
    }

    foreach ($file in $dateFiles) {
        $rel = $file.FullName.Substring($VaultPath.Length).TrimStart('\', '/')
        if ($WhatIf) {
            Write-Log "WHATIF: Would delete $rel"
        } else {
            Remove-Item -LiteralPath $file.FullName -Force
            Write-Log "Deleted: $rel"
        }
    }
    $consolidatedDates.Add($targetDate)
}

if ($totalSessionsFound -eq 0) {
    Write-Log "No individual session files found in the past $LookbackDays day(s) -- nothing to consolidate."
    Write-Log "===== Session Consolidator finished ====="
    exit 0
}

# Combined body for Step 4 doc-update identification
$combinedBody = $allSessionBodies -join "`n`n---`n`n"

# =============================================================================
# STEP 3b -- Upsert All-Sessions.md with a link for each newly consolidated date
# =============================================================================

$allSessionsPath = Join-Path $SessionsDir "All-Sessions.md"

if (-not $WhatIf -and $consolidatedDates.Count -gt 0) {
    $allContent = if (Test-Path -LiteralPath $allSessionsPath) {
        Read-TextFile $allSessionsPath
    } else {
        "---`ndate: `"$Today`"`ntags: [`"sessions`", `"archive`"]`n---`n`n# All Sessions`n"
    }
    foreach ($date in $consolidatedDates) {
        if ($allContent -notmatch [regex]::Escape($date)) {
            $allContent = $allContent.TrimEnd() + "`n`n## $date`n`n- [[$date]] -- Session log for $date"
        }
    }
    Write-TextFile $allSessionsPath $allContent
    Write-Log "Updated All-Sessions.md ($($consolidatedDates.Count) date(s) added)."
}

# =============================================================================
# STEP 4 -- Identify vault documents that need updating
# =============================================================================

$excludeFolders = @("sessions")

$allDocs = Get-ChildItem -LiteralPath $VaultPath -Filter "*.md" -Recurse -File | Where-Object {
    $rel       = $_.FullName.Substring($VaultPath.Length).TrimStart('\', '/')
    $topFolder = ($rel -split '[\\/]')[0]
    $inExcluded = $rel -match '[\\/]' -and $excludeFolders -contains $topFolder.ToLowerInvariant()
    (-not $inExcluded) -and ($_.Name -ne "_Index.md")
}

$docSummaries = foreach ($doc in $allDocs) {
    $rel     = $doc.FullName.Substring($VaultPath.Length).TrimStart('\', '/')
    $raw     = Read-TextFile $doc.FullName
    $body    = $raw -replace '(?s)\A---\r?\n.*?\r?\n---\r?\n', ''
    $excerpt = $body.Trim()
    if ($excerpt.Length -gt 350) { $excerpt = $excerpt.Substring(0, 350) + "..." }
    "PATH: $rel`nEXCERPT: $excerpt"
}

$docSummariesText = $docSummaries -join "`n`n---`n`n"

$identifyPrompt = @"
You are reviewing today's session notes ($Today) for an Obsidian knowledge vault.

TODAY'S SESSIONS:
$combinedBody

EXISTING VAULT DOCUMENTS:
$docSummariesText

Based on what happened today, which existing vault documents (if any) contain information that is now outdated or should be expanded with today's findings?

Respond with ONLY valid JSON in this exact format -- no extra text, no markdown fences:
{"documents_to_update":[{"path":"relative/path.md","what_to_incorporate":"Specific description of what to add or correct"}]}

Rules:
- Only include documents with genuinely meaningful updates (new config, new findings, corrected info)
- Maximum $MaxDocUpdates documents
- Use the exact path as shown under PATH: in EXISTING VAULT DOCUMENTS
- If nothing needs updating, return: {"documents_to_update":[]}
"@

Write-Log "Asking Claude which documents need updating..."
$docsToUpdate = @()

# W1: Skip API call entirely in WhatIf mode
if ($WhatIf) {
    Write-Log "WHATIF: Skipping Claude identification call."
} else {
    try {
        $identifyRaw = Invoke-Claude -Prompt $identifyPrompt `
            -SystemMsg "You are a knowledge base curator. Respond only with valid JSON, no extra text."
        $jsonMatch = [regex]::Match($identifyRaw, '(?s)\{.*\}')
        if ($jsonMatch.Success) {
            $parsed      = $jsonMatch.Value | ConvertFrom-Json
            $docsToUpdate = @($parsed.documents_to_update)
        }
        Write-Log "Claude identified $($docsToUpdate.Count) document(s) to update."
    } catch {
        Write-Log "Could not determine documents to update: $($_.Exception.Message)" "WARN"
    }
}

# =============================================================================
# STEP 5 -- Rewrite each identified document
# =============================================================================

$updated = 0
$failed  = 0

foreach ($docUpdate in $docsToUpdate) {
    $docPath  = $docUpdate.path
    $whatToAdd = $docUpdate.what_to_incorporate
    $fullPath  = Join-Path $VaultPath $docPath

    if (-not (Test-Path -LiteralPath $fullPath)) {
        Write-Log "Document not found, skipping: $docPath" "WARN"
        continue
    }

    try {
        $currentContent = Read-TextFile $fullPath

        $rewritePrompt = @"
You are updating an Obsidian vault note with new information from today's session ($Today).

CURRENT DOCUMENT ($docPath):
$currentContent

NEW INFORMATION TO INCORPORATE:
$whatToAdd

Rewrite the document incorporating this new information. Rules:
1. Preserve ALL existing content -- do not remove or shorten anything
2. Keep YAML frontmatter, tags, and Related Topics section intact
3. Add new information where it logically fits in the existing structure
4. If the document has a date field in frontmatter, update it to $Today
5. Keep the same writing style and heading structure
6. Output ONLY the complete updated markdown -- no preamble or explanation
"@

        Write-Log "Updating: $docPath"

        if ($WhatIf) {
            Write-Log "WHATIF: Would rewrite $docPath"
            $updated++
            continue
        }

        $rewritten = Invoke-Claude -Prompt $rewritePrompt `
            -SystemMsg "You are a technical writer updating a knowledge base document. Output only the complete updated markdown."

        if ([string]::IsNullOrWhiteSpace($rewritten)) {
            throw "Empty rewrite returned."
        }

        # C3: Strip accidental markdown fences the model may add
        $rewritten = $rewritten -replace '(?s)^\s*```(?:markdown|md)?\s*\r?\n', '' `
                                 -replace '(?s)\r?\n```\s*$', ''

        # C3: Sanity size guard -- refuse a rewrite suspiciously smaller than the original
        if ($rewritten.Length -lt ($currentContent.Length * 0.5)) {
            throw "Rewrite suspiciously short ($($rewritten.Length) vs $($currentContent.Length) chars) -- skipping to avoid data loss."
        }

        # C3: Back up original before overwriting
        $backupPath = "$fullPath.bak"
        Copy-Item -LiteralPath $fullPath -Destination $backupPath -Force
        Write-TextFile $fullPath $rewritten
        Write-Log "Updated: $docPath (backup: $($backupPath | Split-Path -Leaf))"
        $updated++
    }
    catch {
        Write-Log "ERROR updating $docPath`: $($_.Exception.Message)" "ERROR"
        $failed++
    }
}

# =============================================================================
# STEP 6 -- Toast notification
# =============================================================================

$summary = "Merged $totalSessionsFound session(s) across $($consolidatedDates.Count) date(s). Updated $updated doc(s)."
if ($failed -gt 0) { $summary += " $failed failed." }
Write-Log $summary

try {
    if (-not $WhatIf) {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        $tmpl = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
            [Windows.UI.Notifications.ToastTemplateType]::ToastText02)
        $tmpl.SelectSingleNode('//text[@id="1"]').InnerText = "Obsidian Session Consolidator"
        $tmpl.SelectSingleNode('//text[@id="2"]').InnerText = $summary
        $toast = [Windows.UI.Notifications.ToastNotification]::new($tmpl)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier(
            "Obsidian Organizer").Show($toast)
    }
} catch { }

Write-Log "===== Session Consolidator finished ====="
