<#
.SYNOPSIS
    Build the add-in and create (or update) a GitHub release containing the built .xlam.

.DESCRIPTION
    Runs `tools/build_and_test.ps1` to produce an .xlam in `dist/`, prepares a release
    artifact named `bysio-addon-<tag>.xlam` and uses the `gh` CLI to create or update
    a GitHub release for the specified tag.

    Requires:
      - Windows with Excel (for the build script)
      - `gh` CLI installed and authenticated
      - `git` configured with a remote `origin` or `-Repo` provided

.EXAMPLE
    .\tools\release.ps1 -Tag v1.0.0 -CreateTag -PushTag

    Creates a local tag `v1.0.0`, pushes it, runs the build, and creates a release
    (or uploads the asset if the release already exists).
#>

Param(
    [Parameter(Mandatory=$false)] [string]$Tag,
    [Parameter(Mandatory=$false)] [string]$Repo,
    [Parameter(Mandatory=$false)] [string]$Title = '',
    [Parameter(Mandatory=$false)] [string]$Notes = '',
    [switch]$CreateTag,
    [switch]$PushTag,
    [switch]$Prerelease,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = $MyInvocation.MyCommand.Definition
$toolsDir = Split-Path -Parent $scriptPath
$repoRoot = Split-Path -Parent $toolsDir
$buildScript = Join-Path $toolsDir 'build_and_test.ps1'
$dist = Join-Path $repoRoot 'dist'

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Error "gh CLI not found. Install from https://cli.github.com/ and authenticate (`gh auth login`)."
    exit 1
}

# Determine repository owner/name if not provided
if (-not $Repo) {
    $remoteUrl = (git config --get remote.origin.url) -as [string]
    if (-not $remoteUrl) {
        try {
            $Repo = gh repo view --json nameWithOwner --jq .nameWithOwner
        }
        catch {
            Write-Error "Could not determine repository. Provide -Repo 'owner/repo' or ensure git remote origin exists or gh is configured."
            exit 1
        }
    }
    else {
        if ($remoteUrl -match '[:/](?<repo>[^/]+/[^/]+)(?:\.git)?$') {
            $Repo = $matches['repo']
        }
        else {
            Write-Error "Unable to parse remote origin url: $remoteUrl"
            exit 1
        }
    }
}

if (-not $Tag) {
    $Tag = Read-Host "Enter tag to release (e.g. v1.0.0)"
    if (-not $Tag) { Write-Error "Tag is required"; exit 1 }
}

if (-not $Title) { $Title = "Release $Tag" }

Write-Host "Repository: $Repo"
Write-Host "Tag: $Tag"
Write-Host "Build script: $buildScript"

if (-not (Test-Path $buildScript)) {
    Write-Error "Build script not found at $buildScript"
    exit 1
}

if ($DryRun) {
    Write-Host "[dry-run] Would run build script: $buildScript"
}
else {
    Write-Host "Running build script..."
    & $buildScript
}

if (-not (Test-Path $dist)) { Write-Error "Dist folder not found: $dist"; exit 1 }

$xlam = Get-ChildItem -Path $dist -Filter '*.xlam' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $xlam) { Write-Error "No .xlam found in $dist"; exit 1 }

$artifactName = "bysio-addon-$Tag.xlam"
$artifactPath = Join-Path $dist $artifactName

if ($DryRun) { Write-Host "[dry-run] Would copy $($xlam.FullName) -> $artifactPath" }
else { Copy-Item $xlam.FullName -Destination $artifactPath -Force; Write-Host "Prepared artifact: $artifactPath" }

# Optionally create and push a git tag
if ($CreateTag) {
    Write-Host "Creating local tag: $Tag"
    if ($DryRun) { Write-Host "[dry-run] git tag -a $Tag -m 'Release $Tag'" }
    else {
        $existing = git tag -l $Tag
        if (-not $existing) { git tag -a $Tag -m "Release $Tag"; Write-Host "Tag created: $Tag" }
        else { Write-Host "Tag already exists locally: $Tag" }
        if ($PushTag) { Write-Host "Pushing tag to origin..."; git push origin $Tag }
    }
}

# Create or update GitHub release
$releaseExists = $false
try { gh release view $Tag -R $Repo > $null 2>&1; if ($LASTEXITCODE -eq 0) { $releaseExists = $true } } catch {}

if ($DryRun) {
    if ($releaseExists) { Write-Host "[dry-run] gh release upload $Tag $artifactPath --clobber --repo $Repo" }
    else {
        $prText = if ($Prerelease) { '--prerelease' } else { '' }
        Write-Host "[dry-run] gh release create $Tag $artifactPath -t '$Title' -n '$Notes' $prText --repo $Repo"
    }
}
else {
    if ($releaseExists) {
        Write-Host "Release $Tag exists — uploading asset (clobber)..."
        gh release upload $Tag $artifactPath --clobber --repo $Repo
    }
    else {
        Write-Host "Creating release $Tag..."
        $prArgs = @()
        if ($Prerelease) { $prArgs += '--prerelease' }
        $args = @($Tag, $artifactPath, '-t', $Title, '-n', $Notes) + $prArgs + @('--repo', $Repo)
        gh release create @args
    }
    if ($LASTEXITCODE -ne 0) { Write-Error "gh release command failed (exit $LASTEXITCODE)"; exit $LASTEXITCODE }
    Write-Host "Release created/updated successfully."
}
