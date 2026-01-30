Param(
  [string]$Configuration = "Release",
  [string]$Platform = "x64",
  [string]$Version = "1.0"
)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$slndir = Join-Path $here 'vs'
$pkgdir = Join-Path $here 'package'
$wdxOut = Join-Path $slndir "$Platform\$Configuration\MsgMetaWDX.wdx"

Write-Host "Building MsgMetaWDX ($Platform/$Configuration) ..."
$msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
  $msbuild = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
}
if (-not (Test-Path $msbuild)) { throw "MSBuild not found. Open Visual Studio and build manually." }

& $msbuild (Join-Path $slndir 'MsgMetaWDX.vcxproj') /p:Configuration=$Configuration /p:Platform=$Platform /t:Build /nologo /v:m
if (-not (Test-Path $wdxOut)) { throw "Build output not found: $wdxOut" }

# Create auto-install zip
$zipName = Join-Path $here ("MsgMetaWDX_v${Version}_" + $Platform + ".zip")
if (Test-Path $zipName) { Remove-Item $zipName }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($pkgdir, $zipName)
# Reopen and add the wdx
$fs = New-Object IO.FileStream($zipName, [IO.FileMode]::Open)
$zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Update)
$entry = $zip.CreateEntry("MsgMetaWDX.wdx")
$es = $entry.Open()
[byte[]]$bytes = [System.IO.File]::ReadAllBytes($wdxOut)
$es.Write($bytes,0,$bytes.Length)
$es.Close()
$zip.Dispose()
$fs.Close()

Write-Host "Created package: $zipName"
