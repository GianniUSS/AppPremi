param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$ExePath,

    [string]$ReleaseNotesPath,

    [string]$ManifestPath = "update_manifest.json",

    [string]$Repository = "GianniUSS/AppPremi"
)

$ErrorActionPreference = "Stop"

function Set-FileUtf8NoBom {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Assert-Tool {
    param(
        [string]$Command,
        [string]$FriendlyName
    )

    if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
        throw "Lo strumento '$FriendlyName' ('$Command') non è disponibile nel PATH."
    }
}

Assert-Tool -Command git -FriendlyName "Git"
Assert-Tool -Command gh -FriendlyName "GitHub CLI"
Assert-Tool -Command certutil -FriendlyName "CertUtil"

if (-not (Test-Path $ExePath)) {
    throw "File eseguibile non trovato: $ExePath"
}

if (-not (Test-Path $ManifestPath)) {
    throw "Manifesto aggiornamenti non trovato: $ManifestPath"
}

if ($ReleaseNotesPath -and -not (Test-Path $ReleaseNotesPath)) {
    throw "File note di rilascio non trovato: $ReleaseNotesPath"
}

$exeFileName = [System.IO.Path]::GetFileName($ExePath)

Write-Host "Calcolo hash SHA256..." -ForegroundColor Cyan
$hashOutput = certutil -hashfile $ExePath SHA256
$hash = ($hashOutput | Select-String -Pattern '^[0-9A-Fa-f]{64}$').Line
if (-not $hash) {
    throw "Impossibile estrarre l'hash SHA256 dall'output di certutil."
}
$hash = $hash.ToLowerInvariant()

Write-Host "Aggiornamento di version.py..." -ForegroundColor Cyan
$versionFile = "version.py"
if (-not (Test-Path $versionFile)) {
    throw "File version.py non trovato."
}
$newVersionContent = ((Get-Content $versionFile) -join "`n") -replace 'APP_VERSION\s*=\s*".*"', "APP_VERSION = `"$Version`""
Set-FileUtf8NoBom -Path $versionFile -Content ($newVersionContent + "`n")

Write-Host "Aggiornamento del manifesto $ManifestPath..." -ForegroundColor Cyan
$manifest = Get-Content $ManifestPath | ConvertFrom-Json
$manifest.latest_version = $Version
$manifest.release_date = (Get-Date -Format 'yyyy-MM-dd')
$manifest.sha256 = $hash
$manifest.file_name = $exeFileName

if ($ReleaseNotesPath) {
    $notes = Get-Content $ReleaseNotesPath | Where-Object { $_.Trim() -ne '' }
    $manifest.release_notes = @($notes)
}

$manifest.download_url = "https://github.com/$Repository/releases/download/v$Version/$exeFileName"
$manifestJson = (ConvertTo-Json $manifest -Depth 5)
Set-FileUtf8NoBom -Path $ManifestPath -Content ($manifestJson + "`n")

Write-Host "Eseguo commit e push..." -ForegroundColor Cyan
git add $versionFile $ManifestPath
git commit -m "chore: release $Version"

git tag "v$Version" -f

git push origin main
git push origin "v$Version" --force

Write-Host "Creo la release GitHub e carico l'asset..." -ForegroundColor Cyan
$ghArgs = @(
    "release", "create", "v$Version", $ExePath,
    "--title", "v$Version",
    "--verify-tag",
    "--latest"
)

if ($ReleaseNotesPath) {
    $ghArgs += @("--notes-file", $ReleaseNotesPath)
} else {
    $ghArgs += @("--notes", "Release $Version")
}

gh @ghArgs

Write-Host "Operazione completata. Manifest e versione aggiornati." -ForegroundColor Green
