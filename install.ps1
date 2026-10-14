# ติดตั้ง skills ทั้งหมดจาก repo นี้ไปไว้ที่ ~/.claude/skills (และ commands ไปที่ ~/.claude/commands)
#
# ใช้งาน:
#   .\install.ps1                     # รันจากโฟลเดอร์ repo ที่ clone มาแล้ว
#   .\install.ps1 -SkipInProgress     # ไม่ติดตั้ง skill กลุ่มทดลอง (in-progress)
#   irm https://raw.githubusercontent.com/jatupornrmutsb/claude-skills/master/install.ps1 | iex

param(
    [switch]$SkipInProgress
)

$ErrorActionPreference = 'Stop'

$RepoUrl = 'https://github.com/jatupornrmutsb/claude-skills.git'
$Dest    = Join-Path $HOME '.claude\skills'
$CmdDest = Join-Path $HOME '.claude\commands'

# ── หาโฟลเดอร์ต้นทาง: ใช้ repo ที่สคริปต์อยู่ ถ้าไม่มีให้ clone ชั่วคราว ─────
$Src = $null
$Tmp = $null
if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot 'skills'))) {
    $Src = $PSScriptRoot
} else {
    $Tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("claude-skills-" + [guid]::NewGuid())
    Write-Host "กำลัง clone $RepoUrl ..."
    git clone --depth 1 --quiet $RepoUrl $Tmp
    if ($LASTEXITCODE -ne 0) { throw 'clone ไม่สำเร็จ' }
    $Src = $Tmp
}

try {
    New-Item -ItemType Directory -Force -Path $Dest, $CmdDest | Out-Null

    # ── รวบรวม skill: skills/<หมวด>/<ชื่อ>/ และ skill ของเราเองที่ root ──
    $skillDirs = @(Get-ChildItem (Join-Path $Src 'skills') -Directory |
        Where-Object { -not ($SkipInProgress -and $_.Name -eq 'in-progress') } |
        ForEach-Object { Get-ChildItem $_.FullName -Directory })
    $skillDirs += @(Get-ChildItem $Src -Directory | Where-Object { $_.Name -ne 'skills' })
    $skillDirs = $skillDirs | Where-Object { Test-Path (Join-Path $_.FullName 'SKILL.md') }

    $count = 0
    foreach ($dir in $skillDirs) {
        $target = Join-Path $Dest $dir.Name
        if (Test-Path $target) { Remove-Item -Recurse -Force $target }
        Copy-Item -Recurse $dir.FullName $target
        Write-Host "  ติดตั้ง $($dir.Name)"
        $count++
    }

    # ── คัดลอก slash commands ─────────────────────────────────────────────
    $cmdSrc = Join-Path $Src 'commands'
    if (Test-Path $cmdSrc) {
        Copy-Item (Join-Path $cmdSrc '*.md') $CmdDest -Force
    }

    Write-Host "เสร็จแล้ว ติดตั้ง $count skill ไปที่ $Dest"
    Write-Host 'เปิด Claude Code ใหม่ (หรือเริ่ม session ใหม่) เพื่อให้เห็น skill'
} finally {
    if ($Tmp -and (Test-Path $Tmp)) { Remove-Item -Recurse -Force $Tmp }
}
