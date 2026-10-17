# 构建辅助：本机没有 Visual Studio / SDK 时也能编出 .exe（原生 .ocx 由 GitHub Actions 产出）
$ErrorActionPreference = 'Stop'
$repo = $PSScriptRoot
$dist = Join-Path $repo 'dist'
New-Item -ItemType Directory -Force -Path $dist | Out-Null

$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$x86 = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$src = @('PptTouchClick.cs','Slideshow.cs','Hook.cs','Tray.cs','TouchInjector.cs') | ForEach-Object { Join-Path $repo $_ }
$ref = '/r:System.dll','/r:System.Core.dll','/r:System.Drawing.dll','/r:System.Windows.Forms.dll'

& $csc /nologo /target:winexe /optimize+ /utf8output /out:"$dist\PptTouchClick.exe" @ref @src
if ($LASTEXITCODE -ne 0) { throw "csc failed: $LASTEXITCODE" }

if (Test-Path $x86) {
    & $x86 /nologo /target:library /optimize+ /utf8output /out:"$dist\PptTouchClick32.dll" @ref @src
}

Copy-Item (Join-Path $repo 'README.md'), (Join-Path $repo 'install.cmd'), (Join-Path $repo 'uninstall.cmd') $dist -Force
Get-ChildItem $dist | Format-Table Name, Length
Write-Host "OK -> $dist"