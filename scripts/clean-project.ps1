$ErrorActionPreference = 'Stop'
$root = "e:\\Works\\Excel-Tools-Pro"
$reportDir = Join-Path $root "Reports"
$reportPath = Join-Path $reportDir ("cleanup-report-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + ".txt")

# Ensure report directory exists
New-Item -ItemType Directory -Force -Path $reportDir | Out-Null

function Write-Report($msg) {
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "[$ts] $msg"
  Write-Host $line
  Add-Content -Path $reportPath -Value $line
}

Write-Report "== 清理开始，根目录: $root =="

# Patterns to remove
$dirPatterns = @('bin','obj','Debug','Release','.vs','.vscode','.idea','.fake','paket-files','node_modules')
$filePatterns = @('*.tmp','*.temp','*.log','*.bak','*.backup','*.old','*.rsuser','*.suo','*.user','*.DotSettings.user','Thumbs.db','.DS_Store')

# Remove directories
foreach ($dp in $dirPatterns) {
  $dirs = Get-ChildItem -Path $root -Recurse -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $dp }
  foreach ($d in $dirs) {
    Write-Report "删除目录: $($d.FullName)"
    try { Remove-Item -Recurse -Force -ErrorAction Stop -LiteralPath $d.FullName } catch { Write-Report "跳过(错误): $($_.Exception.Message)" }
  }
}

# Remove files
foreach ($fp in $filePatterns) {
  $files = Get-ChildItem -Path $root -Recurse -File -Force -ErrorAction SilentlyContinue -Filter $fp
  foreach ($f in $files) {
    Write-Report "删除文件: $($f.FullName)"
    try { Remove-Item -Force -ErrorAction Stop -LiteralPath $f.FullName } catch { Write-Report "跳过(错误): $($_.Exception.Message)" }
  }
}

# Remove test artifacts from .gitignore list
$testArtifacts = @('TestFiles','TestReport.md','TestScript.ps1','VerifyResults.ps1')
foreach ($ta in $testArtifacts) {
  $paths = Get-ChildItem -Path $root -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $ta }
  foreach ($p in $paths) {
    Write-Report "删除测试产物: $($p.FullName)"
    try { Remove-Item -Recurse -Force -ErrorAction Stop -LiteralPath $p.FullName } catch { Write-Report "跳过(错误): $($_.Exception.Message)" }
  }
}

# Prune empty directories
$emptyDirs = Get-ChildItem -Path $root -Directory -Recurse -Force | Where-Object { @(Get-ChildItem -Force -LiteralPath $_.FullName).Count -eq 0 }
foreach ($ed in $emptyDirs) {
  Write-Report "删除空目录: $($ed.FullName)"
  try { Remove-Item -Force -ErrorAction Stop -LiteralPath $ed.FullName } catch { Write-Report "跳过(错误): $($_.Exception.Message)" }
}

Write-Report "== 清理结束 =="
Write-Report "报告路径: $reportPath"