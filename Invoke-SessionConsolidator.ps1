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
# STEP 1 -- Find today's individual session files
# =============================================================================

$SessionsDir = Join-Path $VaultPath "sessions"
if (-not (Test-Path -LiteralPath $SessionsDir)) {
    Write-Log "Sessions folder not found: $SessionsDir" "WARN"
    Write-Log "===== Session Consolidator finished ====="
    exit 0
}

# Individual session files: YYYY-MM-DD-SomeName.md  (date + hyphen + title)
# Exclude: YYYY-MM-DD.md (already-consolidated daily files) and All-Sessions.md
$todayFiles = Get-ChildItem -LiteralPath $SessionsDir -Filter "*.md" -File |
    Where-Object { $_.Name -match "^$Today-.+\.md$" } |
    Sort-Object Name

if ($todayFiles.Count -eq 0) {
    Write-Log "No individual session files found for $Today -- nothing to consolidate."
    Write-Log "===== Session Consolidator finished ====="
    exit 0
}

Write-Log "Found $($todayFiles.Count) session file(s) for $Today."

# =============================================================================
# STEP 2 -- Read and merge sessions
# =============================================================================

$sessionParts = foreach ($file in $todayFiles) {
    $raw   = Read-TextFile $file.FullName
    # Strip YAML frontmatter block (anchored to start of file)
    $body  = $raw -replace '(?s)\A---\r?\n.*?\r?\n---\r?\n', ''
    $title = $file.BaseName -replace "^$Today-", '' -replace '-', ' '
    "### $title`n`n$($body.Trim())"
}

$combinedBody = $sessionParts -join "`n`n---`n`n"

$consolidatedContent = @"
---
date: "$Today"
type: "daily-session"
tags: ["session", "daily-log", "category/sessions"]
---

#session #daily-log #category/sessions

# Session Log -- $Today

$combinedBody

## Related Topics

- [[All-Sessions]]
"@

# =============================================================================
# STEP 3 -- Write consolidated file, delete individual files
# =============================================================================

$consolidatedPath = Join-Path $SessionsDir "$Today.md"
$consolidatedRel  = "sessions\$Today.md"

if ($WhatIf) {
    Write-Log "WHATIF: Would write $consolidatedRel"
} else {
    Write-TextFile $consolidatedPath $consolidatedContent
    Write-Log "Wrote: $consolidatedRel"
}

foreach ($file in $todayFiles) {
    $rel = $file.FullName.Substring($VaultPath.Length).TrimStart('\', '/')
    if ($WhatIf) {
        Write-Log "WHATIF: Would delete $rel"
    } else {
        Remove-Item -LiteralPath $file.FullName -Force
        Write-Log "Deleted: $rel"
    }
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

$summary = "Merged $($todayFiles.Count) session(s) into $Today.md. Updated $updated doc(s)."
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
