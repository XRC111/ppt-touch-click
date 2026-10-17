# 生成 COM 注册用类型库（src\PptSlideProc.tlb）
# 需要 Windows SDK 的 tlbexp.exe；没有 tlbexp 时用本地的 regasm.exe（需要能写 HKCU\Software\Classes，CI 上可行）。
$ErrorActionPreference = 'Stop'
$repo = $PSScriptRoot
$out = Join-Path $repo 'src\PptSlideProc.tlb'

$sdk = 'C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools'
$tlbexp = Get-ChildItem $sdk -Recurse -Filter tlbexp.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($tlbexp) {
    & $tlbexp.FullName (Join-Path $repo 'src\PptSlideProc.TlbGen.dll') /out:$out
} else {
    & (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\regasm.exe') /tlb:$out (Join-Path $repo 'src\PptSlideProc.TlbGen.dll')
}
if (-not (Test-Path $out)) { throw "类型库生成失败：$out" }
Get-Item $out | Format-List Name, Length