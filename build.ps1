<#
  neisXLS 빌드 스크립트

  template.xlam 을 복사한 뒤 src 폴더의 VBA 모듈을 모두 다시 넣어
  build\neisXLS.xlam 을 만듭니다.
#>
[CmdletBinding()]
param(
    [string]$Output
)

$ErrorActionPreference = 'Stop'
$repo = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Output)) {
    $Output = Join-Path $repo 'neisXLS.xlam'
}
$src = Join-Path $repo 'src'
$srcEnc = Join-Path $repo 'build\src_cp949'
$template = Join-Path $repo 'template.xlam'

if (-not (Test-Path -LiteralPath $template)) {
    throw "template.xlam 이 없습니다: $template"
}
if (-not (Test-Path -LiteralPath $src)) {
    throw "src 폴더가 없습니다: $src"
}

$outDir = Split-Path -Parent $Output
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

if (Test-Path -LiteralPath $srcEnc) {
    Remove-Item -LiteralPath $srcEnc -Recurse -Force
}
New-Item -ItemType Directory -Path $srcEnc -Force | Out-Null

$utf8 = New-Object System.Text.UTF8Encoding($false)
$cp949 = [System.Text.Encoding]::GetEncoding(949)

foreach ($file in Get-ChildItem -LiteralPath $src -File) {
    $targetFile = Join-Path $srcEnc $file.Name
    if ($file.Extension -eq '.frx') {
        Copy-Item -LiteralPath $file.FullName -Destination $targetFile -Force
        continue
    }
    if ($file.Extension -notin '.bas', '.cls', '.frm') {
        continue
    }

    $text = [IO.File]::ReadAllText($file.FullName, $utf8)
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`n", "`r`n"
    [IO.File]::WriteAllText($targetFile, $text, $cp949)
}

Copy-Item -LiteralPath $template -Destination $Output -Force
Unblock-File -LiteralPath $Output -ErrorAction SilentlyContinue

$xl = $null
$wb = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 3

    $wb = $xl.Workbooks.Open($Output, 0, $false)

    Write-Host '기존 VBA 모듈 제거...'
    $components = @($wb.VBProject.VBComponents)
    foreach ($component in $components) {
        if ($component.Type -ne 100) {
            $wb.VBProject.VBComponents.Remove($component)
        }
    }

    Write-Host 'src 모듈 가져오기...'
    $files = Get-ChildItem -LiteralPath $srcEnc -File |
        Where-Object {
            $_.Extension -in '.bas', '.cls', '.frm' -and
            $_.Name -notin @('Sheet1.cls', '현재_통합_문서.cls')
        } |
        Sort-Object Name

    foreach ($file in $files) {
        $wb.VBProject.VBComponents.Import($file.FullName)
    }

    $wb.Save()
    $wb.Close($true)
    $wb = $null

    Write-Host "빌드 완료: $Output"
}
catch {
    Write-Error $_
    if ($wb) {
        try { $wb.Close($false) } catch {}
    }
    throw
}
finally {
    if ($xl) {
        try { $xl.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
