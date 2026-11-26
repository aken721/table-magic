$ErrorActionPreference = "Stop"

$projectDir = "TableMagic.Cli"

Write-Host "=== TableMagic Skill Package Builder ===" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/4] Building single-file self-contained (win-x64)..." -ForegroundColor Yellow
$scDir = "publish-sc"
if (Test-Path $scDir) { Remove-Item $scDir -Recurse -Force }
dotnet publish $projectDir -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=none -p:DebugSymbols=false -o $scDir
if ($LASTEXITCODE -ne 0) { Write-Error "Build failed!"; exit 1 }

Write-Host "[2/4] Staging files..." -ForegroundColor Yellow
$stagingDir = "staging-tablemagic"
if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }
New-Item -ItemType Directory -Path $stagingDir | Out-Null

Copy-Item -Path "$scDir\tablemagic.exe" -Destination $stagingDir

$nativeDir = Join-Path $stagingDir "runtimes"
New-Item -ItemType Directory -Path $nativeDir -Force | Out-Null
New-Item -ItemType Directory -Path "$nativeDir\win-x64\native" -Force | Out-Null

if (Test-Path "$scDir\runtimes\win-x64\native") {
    Get-ChildItem "$scDir\runtimes\win-x64\native" -File | Where-Object {
        $_.Extension -notin '.pdb'
    } | ForEach-Object {
        Copy-Item $_.FullName -Destination "$nativeDir\win-x64\native"
    }
}

Copy-Item -Path "$projectDir\skill.md" -Destination $stagingDir
Copy-Item -Path "$projectDir\mcp-config.json" -Destination $stagingDir
Copy-Item -Path "$projectDir\README.md" -Destination $stagingDir

Write-Host "[3/4] Creating ZIP..." -ForegroundColor Yellow
if (Test-Path "table-magic-skill.zip") { Remove-Item "table-magic-skill.zip" }
Compress-Archive -Path "$stagingDir\*" -DestinationPath "table-magic-skill.zip" -Force
Remove-Item $stagingDir -Recurse -Force

$scSize = [math]::Round((Get-Item "table-magic-skill.zip").Length / 1MB, 2)
Write-Host "  -> table-magic-skill.zip ($scSize MB) - single-file, self-contained" -ForegroundColor Green

Write-Host ""
Write-Host "[4/4] Building NuGet tool package..." -ForegroundColor Yellow
dotnet pack $projectDir -c Release -o ./nupkg
if ($LASTEXITCODE -ne 0) { Write-Error "Pack failed!"; exit 1 }
Write-Host "  -> nupkg/TableMagic.Cli.2.5.2.1.nupkg" -ForegroundColor Green

Write-Host ""
Write-Host "Cleanup..." -ForegroundColor Yellow
if (Test-Path $scDir) { Remove-Item $scDir -Recurse -Force }
if (Test-Path "publish_singlefile") { Remove-Item "publish_singlefile" -Recurse -Force }
if (Test-Path "publish_fdd") { Remove-Item "publish_fdd" -Recurse -Force }

Write-Host ""
Write-Host "=== Done! ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Installation methods:" -ForegroundColor White
Write-Host ""
Write-Host "Method 1: Extract ZIP (simplest)" -ForegroundColor Yellow
Write-Host "  1. Extract table-magic-skill.zip to a directory"
Write-Host "  2. Add to Agent MCP config:"
Write-Host '     { "mcpServers": { "tablemagic": { "command": "<path>/tablemagic", "args": ["mcp"] } } }'
Write-Host ""
Write-Host "Method 2: dotnet tool (requires .NET 8 SDK)" -ForegroundColor Yellow
Write-Host "  1. dotnet tool install --global --add-source ./nupkg TableMagic.Cli"
Write-Host "  2. Add to Agent MCP config:"
Write-Host '     { "mcpServers": { "tablemagic": { "command": "tablemagic", "args": ["mcp"] } } }'
