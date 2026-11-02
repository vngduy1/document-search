[CmdletBinding()]
param(
    [ValidateSet("onedir", "onefile")]
    [string]$Mode = "onefile"
)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$entryPoint = Join-Path $projectRoot "src\document_search\__main__.py"
$sourcePath = Join-Path $projectRoot "src"
$executable = if ($Mode -eq "onedir") {
    Join-Path $projectRoot "dist\DocumentSearch\DocumentSearch.exe"
} else {
    Join-Path $projectRoot "dist\DocumentSearch.exe"
}

if (-not (Test-Path -LiteralPath $python)) {
    throw "Existing virtual environment was not found: $python"
}

Push-Location $projectRoot
$originalPath = $env:PATH
try {
    $buildPath = foreach ($directory in ($originalPath -split ';')) {
        if (-not $directory -or -not (Test-Path -LiteralPath $directory)) {
            continue
        }
        $icuData = Get-ChildItem -LiteralPath $directory -Filter 'icudt*.dll' -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $icuData) {
            $directory
        }
    }
    $env:PATH = $buildPath -join ';'

    $bundleOption = if ($Mode -eq "onedir") { "--onedir" } else { "--onefile" }
    $consoleOption = if ($Mode -eq "onedir") { "--console" } else { "--windowed" }
    $arguments = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--noupx",
        $bundleOption,
        $consoleOption,
        "--name", "DocumentSearch",
        "--paths", $sourcePath,
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--hidden-import", "win32com.client",
        "--collect-submodules", "win32com",
        $entryPoint
    )
    & $python @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed with exit code $LASTEXITCODE"
    }
    $analysis = Join-Path $projectRoot "build\DocumentSearch\Analysis-00.toc"
    if (Select-String -LiteralPath $analysis -SimpleMatch "icuuc.dll" -Quiet) {
        throw "Build captured a conflicting non-system ICU DLL from PATH"
    }

    & $executable --self-test
    if ($LASTEXITCODE -ne 0) {
        throw "Packaged SQLite FTS5 self-test failed with exit code $LASTEXITCODE"
    }

    & $executable --smoke-test-gui
    if ($LASTEXITCODE -ne 0) {
        throw "Packaged GUI smoke test failed with exit code $LASTEXITCODE"
    }
}
finally {
    $env:PATH = $originalPath
    Pop-Location
}

Write-Output $executable
