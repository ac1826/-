param(
    [string]$RepoDir = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [string]$Branch = 'main',
    [int]$PollSeconds = 3,
    [int]$DebounceSeconds = 5
)

$ErrorActionPreference = 'Continue'
$ModelFileName = 'app_' + [string][char]0x9c9c + [string][char]0x51bb + [string][char]0x54c1 + [string][char]0x5360 + [string][char]0x6bd4 + '.py'
$ModelPath = Join-Path $RepoDir $ModelFileName
$LogDir = Join-Path $RepoDir 'monitor_logs'
$LogFile = Join-Path $LogDir 'model_auto_sync.log'

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -LiteralPath $LogFile -Value "[$stamp] $Message" -Encoding UTF8
}

function Invoke-Git {
    param([string[]]$Args)
    $output = & git @Args 2>&1
    $code = $LASTEXITCODE
    if ($output) {
        Write-Log ($output -join "`n")
    }
    return $code
}

if (!(Test-Path -LiteralPath $ModelPath)) {
    Write-Log "Model file not found: $ModelPath"
    exit 1
}

Write-Log "Started model auto sync for $ModelPath -> origin/$Branch"
$lastWrite = (Get-Item -LiteralPath $ModelPath).LastWriteTimeUtc

while ($true) {
    Start-Sleep -Seconds $PollSeconds
    try {
        $currentWrite = (Get-Item -LiteralPath $ModelPath).LastWriteTimeUtc
        if ($currentWrite -eq $lastWrite) {
            continue
        }

        $lastWrite = $currentWrite
        Start-Sleep -Seconds $DebounceSeconds

        $diffCode = Invoke-Git @('-C', $RepoDir, 'diff', '--quiet', '--', $ModelFileName)
        if ($diffCode -eq 0) {
            Write-Log 'Detected save, but no model diff to sync.'
            continue
        }

        $message = 'Auto-sync model change ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Write-Log "Committing model change: $message"

        $addCode = Invoke-Git @('-C', $RepoDir, 'add', '--', $ModelFileName)
        if ($addCode -ne 0) {
            Write-Log "git add failed with code $addCode"
            continue
        }

        $commitCode = Invoke-Git @('-C', $RepoDir, 'commit', '-m', $message, '--', $ModelFileName)
        if ($commitCode -ne 0) {
            Write-Log "git commit skipped/failed with code $commitCode"
            continue
        }

        $pullCode = Invoke-Git @('-C', $RepoDir, 'pull', '--rebase', 'origin', $Branch)
        if ($pullCode -ne 0) {
            Write-Log "git pull --rebase failed with code $pullCode; push skipped."
            continue
        }

        $pushCode = Invoke-Git @('-C', $RepoDir, 'push', 'origin', $Branch)
        if ($pushCode -ne 0) {
            Write-Log "git push failed with code $pushCode"
            continue
        }

        Write-Log 'Model change synced to GitHub.'
    } catch {
        Write-Log ("Unhandled error: " + $_.Exception.Message)
    }
}
