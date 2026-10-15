# insert.html -> 동봉자료_외부감사준비안내.pdf (Edge headless)
# 사용: powershell -ExecutionPolicy Bypass -File brochure\build_insert.ps1
$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
$src  = "file:///" + (Join-Path $dir "insert.html").Replace('\', '/')
$out  = Join-Path (Split-Path -Parent $dir) "동봉자료_외부감사준비안내.pdf"

$dom = & $edge --headless=new --disable-gpu --virtual-time-budget=10000 --dump-dom $src 2>$null
$ov  = [regex]::Match(($dom -join "`n"), 'data-overflow="([^"]*)"').Groups[1].Value
Write-Host "넘침(px, 쪽별): $ov   (0이면 정상)"

& $edge --headless=new --disable-gpu --no-pdf-header-footer --virtual-time-budget=10000 "--print-to-pdf=$out" $src 2>$null | Out-Null
Write-Host "PDF: $out"
