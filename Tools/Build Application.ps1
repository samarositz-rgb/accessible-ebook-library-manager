$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$source = Join-Path $root 'Source'
$application = Join-Path $root 'Application'
$python = 'C:\Program Files\Python314\python.exe'
if (-not (Test-Path -LiteralPath $python)) { $python = (Get-Command python -ErrorAction Stop).Source }
$running = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith($application + '\', [StringComparison]::OrdinalIgnoreCase) })
if ($running.Count) { throw 'Close Accessible Ebook Library Manager before building a replacement.' }
$env:PATH = ($env:PATH -split ';' | Where-Object { $_ -notmatch 'poppler\\Library\\bin' }) -join ';'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$build = Join-Path $root 'Build'
$stage = Join-Path ([IO.Path]::GetTempPath()) ('aelm-build-' + $stamp)
New-Item -ItemType Directory -Path $stage -Force | Out-Null
Push-Location $source
try {
    & $python -c 'from PySide6.QtCore import Qt; import PyInstaller'
    if ($LASTEXITCODE -ne 0) { throw 'Python needs working PySide6 and PyInstaller installations.' }
    & $python prepare_braille_runtime.py
    if ($LASTEXITCODE -ne 0) { throw 'The bundled braille converter could not be verified.' }
    & $python -m PyInstaller --noconfirm --clean --onedir --windowed --name 'Accessible Ebook Library Manager' --add-data "$source\vendor\braille;vendor\braille" --add-data "$root\Documentation\Braille Export Guide.md;." --distpath $stage --workpath (Join-Path $build 'PyInstaller') --specpath $build qt_library_manager.py
    if ($LASTEXITCODE -ne 0) { throw 'Application build failed. The previous application was preserved.' }
} finally { Pop-Location }
$built = Join-Path $stage 'Accessible Ebook Library Manager'
$exe = Join-Path $built 'Accessible Ebook Library Manager.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw 'The build did not produce an executable.' }
$running = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith($application + '\', [StringComparison]::OrdinalIgnoreCase) })
if ($running.Count) { throw "Close the application. The new build is waiting in $built" }
$archive = Join-Path $root ('Archive\Application ' + $stamp)
if (-not ([IO.Path]::GetFullPath($built)).StartsWith([IO.Path]::GetFullPath($stage) + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected staging path.' }
foreach ($path in @($application, $archive)) {
    if (-not ([IO.Path]::GetFullPath($path)).StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected build destination.' }
}
if (Test-Path -LiteralPath $application) { Move-Item -LiteralPath $application -Destination $archive }
try { Move-Item -LiteralPath $built -Destination $application }
catch {
    if ((Test-Path -LiteralPath $archive) -and -not (Test-Path -LiteralPath $application)) { Move-Item -LiteralPath $archive -Destination $application }
    throw
}
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path $root 'Accessible Ebook Library Manager.lnk'))
$shortcut.TargetPath = Join-Path $application 'Accessible Ebook Library Manager.exe'
$shortcut.WorkingDirectory = $application
$shortcut.Description = 'Open Accessible Ebook Library Manager'
$shortcut.Save()
Write-Output ('Ready: ' + $shortcut.TargetPath)
