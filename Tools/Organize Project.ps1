$ErrorActionPreference = 'Stop'
$workspace = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$old = Join-Path $workspace 'Accessible ebook manager'
$new = Join-Path $workspace 'Accessible Ebook Library Manager'
if (Test-Path -LiteralPath $old) {
    if (Test-Path -LiteralPath $new) {
        $remaining = @(Get-ChildItem -LiteralPath $old -Force)
        if ($remaining.Count -ne 1 -or $remaining[0].Name -ne '.git' -or @(Get-ChildItem -LiteralPath $remaining[0].FullName -Force).Count -ne 0) { throw 'Both project folders contain data; inspect before continuing.' }
        $shellArchive = Join-Path $new 'Archive\Original Folder Shell'
        if (-not ([IO.Path]::GetFullPath($shellArchive)).StartsWith([IO.Path]::GetFullPath($new) + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected archive path.' }
        New-Item -ItemType Directory -Path (Join-Path $new 'Archive') -Force | Out-Null
        Move-Item -LiteralPath $old -Destination $shellArchive -Force
    } else {
    $running = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith($old + '\', [StringComparison]::OrdinalIgnoreCase) })
    if ($running.Count) { throw 'Please close the library manager before organizing the project.' }
    if ((Split-Path ([IO.Path]::GetFullPath($old)) -Parent) -ne $workspace -or (Split-Path ([IO.Path]::GetFullPath($new)) -Parent) -ne $workspace) { throw 'Unexpected project path.' }
    Move-Item -LiteralPath $old -Destination $new -Force
    }
}
$root = (Resolve-Path -LiteralPath $new).Path
$moves = [Collections.Generic.List[object]]::new()
function Move-ProjectItem([string]$from, [string]$to) {
    if (-not (Test-Path -LiteralPath $from)) { return }
    $source = (Resolve-Path -LiteralPath $from).Path
    $target = [IO.Path]::GetFullPath($to)
    if (-not $source.StartsWith($workspace + '\', [StringComparison]::OrdinalIgnoreCase) -or -not $target.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Unsafe move: $source -> $target" }
    if (Test-Path -LiteralPath $target) { throw "Target already exists: $target" }
    Move-Item -LiteralPath $source -Destination $target
    $moves.Add(@{from=$source; to=$target})
}
foreach ($folder in @('Application','Documentation','Source','Tools','Archive','Archive\Legacy Build Scripts','Archive\Build Configurations','Archive\Legacy Application','Archive\Test Libraries')) {
    New-Item -ItemType Directory -Path (Join-Path $root $folder) -Force | Out-Null
}
foreach ($file in Get-ChildItem -LiteralPath $root -File -Filter '*.py') { Move-ProjectItem $file.FullName (Join-Path $root ('Source\' + $file.Name)) }
foreach ($folder in @('tests','vendor')) { Move-ProjectItem (Join-Path $root $folder) (Join-Path $root ('Source\' + $folder)) }
foreach ($file in Get-ChildItem -LiteralPath $root -File -Filter '*.md') { Move-ProjectItem $file.FullName (Join-Path $root ('Documentation\' + $file.Name)) }
Move-ProjectItem (Join-Path $root 'Documentation\BRAILLE_EXPORT.md') (Join-Path $root 'Documentation\Braille Export Guide.md')
Move-ProjectItem (Join-Path $root 'Documentation\M4B_RESEARCH.md') (Join-Path $root 'Documentation\Audio Export Research.md')
Move-ProjectItem (Join-Path $root 'Documentation\CHANGELOG.md') (Join-Path $root 'Documentation\Release Notes.md')
foreach ($file in Get-ChildItem -LiteralPath $root -File -Filter '*.bat') { Move-ProjectItem $file.FullName (Join-Path $root ('Archive\Legacy Build Scripts\' + $file.Name)) }
foreach ($file in Get-ChildItem -LiteralPath $root -File -Filter '*.spec') { Move-ProjectItem $file.FullName (Join-Path $root ('Archive\Build Configurations\' + $file.Name)) }
foreach ($file in Get-ChildItem -LiteralPath $root -File -Filter '*.exe') { Move-ProjectItem $file.FullName (Join-Path $root ('Archive\Legacy Application\' + $file.Name)) }
Move-ProjectItem (Join-Path $root 'build') (Join-Path $root 'Archive\Build Cache')
Move-ProjectItem (Join-Path $root 'dist') (Join-Path $root 'Archive\Previous Builds')
Move-ProjectItem (Join-Path $root '__pycache__') (Join-Path $root 'Archive\Python Cache')
foreach ($folder in @('test_appdata','test_appdata_file_conflict','test_appdata_watched')) { Move-ProjectItem (Join-Path $workspace $folder) (Join-Path $root ('Archive\Test Libraries\' + $folder)) }
$moves | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root ('Archive\Folder Changes ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.json')) -Encoding UTF8
Write-Output "Organized project: $root"
