#!/bin/sh
# 生成 COM 注册用的类型库 src/PptSlideProc.tlb
# 需要 Windows SDK 的 tlbexp.exe；没有 tlbexp 时退回本地 regasm.exe（需要可写 HKCU\Software\Classes，CI 上可行）。
set -e
here=$(cd "$(dirname "$0")" && pwd)
stub="$here/src/PptSlideProc.TlbGen.dll"
out="$here/src/PptSlideProc.tlb"

tlbexp=$(ls "/c/Program Files (x86)/Microsoft SDKs/Windows/v10.0A/bin/NETFX 4.8.1 Tools/tlbexp.exe" 2>/dev/null || true)
if [ -z "$tlbexp" ]; then
  tlbexp=$(find "/c/Program Files (x86)/Microsoft SDKs/Windows" -name tlbexp.exe 2>/dev/null | head -n 1)
fi

if [ -n "$tlbexp" ]; then
  "$tlbexp" "$stub" /out:"$out"
else
  "/c/Windows/Microsoft.NET/Framework64/v4.0.30319/regasm.exe" /tlb:"$out" "$stub"
fi

ls -l "$out"
