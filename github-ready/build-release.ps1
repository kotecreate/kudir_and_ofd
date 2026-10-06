$ErrorActionPreference='Stop'
$version=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'VERSION')).Trim()
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw 'VERSION must contain a version such as 1.0.0' }
$output=Join-Path $PSScriptRoot 'dist'
[void][IO.Directory]::CreateDirectory($output)
# Explicit allowlist: never package input files, reports or Git history.
$names=@('compare.ps1','auto-compare.ps1','Запустить.cmd','Инструкция.html','СНАЧАЛА ПРОЧИТАЙТЕ.txt','README.md','LICENSE','CHANGELOG.md','VERSION','RELEASE_NOTES.md')
$files=@($names | ForEach-Object { Join-Path $PSScriptRoot $_ })
foreach ($file in $files) { if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Missing release file: $file" } }
$zip=Join-Path $output "kudir-ofd-$version.zip"
Compress-Archive -LiteralPath $files -DestinationPath $zip -Force
$hash=(Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText((Join-Path $output "kudir-ofd-$version.sha256"),"$hash  kudir-ofd-$version.zip`n",[Text.UTF8Encoding]::new($false))
Write-Output $zip
