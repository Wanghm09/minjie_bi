param(
  [string]$Message = "chore: sync local prototype update"
)

$ErrorActionPreference = 'Stop'

$srcHtml = 'D:/skilltohtml/0408-skilltohtml-prototype-iceblue-v9.html'
$srcLogo = 'D:/skilltohtml/assets/platform-logo.png'
$repo = 'D:/skilltohtml/minjie_bi'

if (!(Test-Path $repo)) {
  git clone https://github.com/Wanghm09/minjie_bi.git $repo
}

if (!(Test-Path $srcHtml)) { throw "Source html not found: $srcHtml" }

Copy-Item -Path $srcHtml -Destination (Join-Path $repo '0408-skilltohtml-prototype-iceblue-v9.html') -Force
if (Test-Path $srcLogo) {
  if (!(Test-Path (Join-Path $repo 'assets'))) { New-Item -ItemType Directory -Path (Join-Path $repo 'assets') | Out-Null }
  Copy-Item -Path $srcLogo -Destination (Join-Path $repo 'assets/platform-logo.png') -Force
}

Set-Location $repo

git add 0408-skilltohtml-prototype-iceblue-v9.html assets/platform-logo.png
if (-not (git diff --cached --quiet)) {
  git commit -m $Message
  git push
  Write-Output 'Synced and pushed.'
} else {
  Write-Output 'No changes to sync.'
}
