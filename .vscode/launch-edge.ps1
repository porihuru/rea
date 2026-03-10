param(
  [Parameter(Mandatory = $true)]
  [string]$TargetFile
)

$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edgePath)) {
  Write-Error "Edge executable not found: $edgePath"
  exit 1
}

$resolvedTarget = Resolve-Path $TargetFile
$profileDir = Join-Path $PSScriptRoot "edge-profile"
if (-not (Test-Path $profileDir)) {
  New-Item -ItemType Directory -Path $profileDir | Out-Null
}

$args = @(
  "--new-window",
  "--user-data-dir=$profileDir",
  "$resolvedTarget"
)

$proc = Start-Process -FilePath $edgePath -ArgumentList $args -PassThru

try {
  Wait-Process -Id $proc.Id
}
finally {
  if (-not $proc.HasExited) {
    Stop-Process -Id $proc.Id -Force
  }
}
