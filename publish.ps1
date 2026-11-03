<#
  B-XLAM xlam 배포 스크립트

  동작
  1. 이 폴더에서 가장 최근에 만들어진 .xlam 파일을 찾습니다.
  2. %AppData%\Microsoft\AddIns 폴더에 같은 이름으로 복사합니다.
  3. 변경 사항을 git에 커밋하고 origin으로 푸시합니다.

  사용 예
    .\publish.ps1
    .\publish.ps1 -Message "예산보드 기능 수정"
    .\publish.ps1 -SkipPush
#>
[CmdletBinding()]
param(
    [string]$Message,
    [switch]$SkipPush
)

$ErrorActionPreference = 'Stop'
$repo = $PSScriptRoot
$addins = Join-Path $env:APPDATA 'Microsoft\AddIns'

function Write-Step([string]$Text) {
    Write-Host "== $Text" -ForegroundColor Cyan
}

Write-Step 'xlam 파일 확인'
$latest = Get-ChildItem -LiteralPath $repo -Filter '*.xlam' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $latest) {
    throw "이 폴더에 .xlam 파일이 없습니다: $repo"
}

$fs = [IO.File]::OpenRead($latest.FullName)
try {
    $sig = New-Object byte[] 2
    [void]$fs.Read($sig, 0, 2)
}
finally {
    $fs.Dispose()
}

if ($sig[0] -ne 0x50 -or $sig[1] -ne 0x4B) {
    throw "xlam 파일이 올바른 Excel 추가기능(ZIP) 형식이 아닙니다: $($latest.Name)"
}

Write-Host "대상 파일: $($latest.Name) ($([Math]::Round($latest.Length / 1MB, 2)) MB, $($latest.LastWriteTime))"

Write-Step 'AddIns 폴더에 설치'
if (-not (Test-Path -LiteralPath $addins)) {
    New-Item -ItemType Directory -Path $addins -Force | Out-Null
}

$target = Join-Path $addins $latest.Name
$excel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    Write-Warning 'Excel이 실행 중입니다. 추가기능 파일이 잠겨 있으면 복사가 실패할 수 있습니다.'
}

$copyOk = $false
try {
    Copy-Item -LiteralPath $latest.FullName -Destination $target -Force
    $copyOk = $true
    Write-Host "복사 완료: $target"
    Write-Host 'Excel을 다시 시작하면 새 버전이 반영됩니다.'
}
catch {
    Write-Warning "AddIns 복사 실패: $($_.Exception.Message)"
    Write-Warning 'Excel을 모두 종료한 뒤 이 스크립트를 다시 실행하세요.'
}

Write-Step 'GitHub 업로드'
Push-Location $repo
try {
    if (-not (Test-Path -LiteralPath (Join-Path $repo '.git'))) {
        throw 'Git 저장소가 아닙니다. 먼저 git init 과 remote 설정이 필요합니다.'
    }

    $changes = git status --porcelain
    if ($changes) {
        if ([string]::IsNullOrWhiteSpace($Message)) {
            $Message = "Update xlam $($latest.Name) ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
        }

        git add -A
        if ($LASTEXITCODE -ne 0) { throw 'git add 실패' }

        git commit -m $Message
        if ($LASTEXITCODE -ne 0) { throw 'git commit 실패' }

        Write-Host "커밋 완료: $Message"
    }
    else {
        Write-Host '커밋할 변경 사항이 없습니다.'
    }

    if (-not $SkipPush) {
        git push origin HEAD
        if ($LASTEXITCODE -ne 0) { throw 'git push 실패' }
        Write-Host 'GitHub 업로드 완료.'
    }
}
finally {
    Pop-Location
}

if (-not $copyOk) {
    Write-Warning 'GitHub 업로드는 끝났지만 AddIns 설치가 실패했습니다. Excel 종료 후 다시 실행하세요.'
    exit 1
}
