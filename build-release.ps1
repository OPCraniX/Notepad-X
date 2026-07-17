$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = if ($env:PYTHON) { $env:PYTHON } else { 'python' }
$SpecPath = Join-Path $ProjectRoot 'Notepad-X.spec'
$ExePath = Join-Path $ProjectRoot 'dist\Notepad-X.exe'

Push-Location $ProjectRoot
try {
    & $Python -m pip check
    if ($LASTEXITCODE -ne 0) { throw 'Python dependency validation failed.' }

    & $Python -c 'import cryptography, PIL, spellchecker, PyInstaller'
    if ($LASTEXITCODE -ne 0) { throw 'One or more release dependencies are not installed.' }

    & $Python -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) { throw 'Unit tests failed.' }

    & $Python -m PyInstaller --clean --noconfirm $SpecPath
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $ExePath -PathType Leaf)) {
        throw 'PyInstaller did not produce dist\Notepad-X.exe.'
    }

    $Manifest = (& $Python -m PyInstaller.utils.cliutils.archive_viewer -l $ExePath 2>&1) -join "`n"
    # archive_viewer prints Python-style path representations on Windows.
    $Manifest = $Manifest.Replace('\\', '/').Replace('\', '/')
    foreach ($RequiredAsset in @(
        'cfg/language/en_us.yml',
        'cfg/themes/default.json',
        'cfg/spellcheck/en.json.gz',
        'Notepad-X-help.txt'
    )) {
        if ($Manifest -notmatch [regex]::Escape($RequiredAsset)) {
            throw "Packaged asset is missing: $RequiredAsset"
        }
    }

    foreach ($ForbiddenAsset in @(
        '.session.json',
        '.editor.json',
        '.notepadx.notes.json',
        '.notepadx.editors.json',
        'Notepad-X.recovery.json',
        'remote-cache',
        'cfg/backups',
        '.git/'
    )) {
        if ($Manifest -match [regex]::Escape($ForbiddenAsset)) {
            throw "Private or development data was packaged: $ForbiddenAsset"
        }
    }

    $Hash = (Get-FileHash -LiteralPath $ExePath -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -LiteralPath (Join-Path $ProjectRoot 'dist\Notepad-X.exe.sha256') -Value "$Hash  Notepad-X.exe" -Encoding ascii
    Write-Host "Built $ExePath"
    Write-Host "SHA-256: $Hash"
}
finally {
    Pop-Location
}
