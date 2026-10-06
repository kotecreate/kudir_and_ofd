$ErrorActionPreference='Stop'
. "$PSScriptRoot/compare.ps1" -NoGui
$kpath=(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*_2026_1_*.xlsx').FullName
$opath=(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*2026-10-01*.xlsx').FullName
$beforeK=(Get-FileHash -LiteralPath $kpath).Hash; $beforeO=(Get-FileHash -LiteralPath $opath).Hash
if ($env:KUDIR_TEST_CACHE -eq '1') { $inputs=Import-Clixml -LiteralPath (Join-Path $PSScriptRoot 'test-output/real-inputs.xml') }
else {
    $inputs=Read-AutoInputs $kpath $opath
    $inputs | Export-Clixml -LiteralPath (Join-Path $PSScriptRoot 'test-output/real-inputs.xml') -Depth 8
}
$r=Get-AutoComparison $inputs.K $inputs.O
$expected=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-output/independent-real.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ($r.Issues.Count) { $r.Issues | Format-List | Out-String | Write-Output; throw 'Unexpected import issues' }
if ($r.K.Count -ne $expected.kcount -or $r.O.Count -ne $expected.ocount) { throw 'Record count mismatch' }
if ($r.KAll -ne [decimal]$expected.kall -or $r.OSales -ne [decimal]$expected.o_receipts -or $r.KLedger -ne [decimal]$expected.kledger) { throw 'Independent totals mismatch' }
if ($r.KSales -ne [decimal]1006650 -or $r.Difference -ne [decimal]2470.04) { throw 'Comparable sales mismatch' }
$diff=@($r.Bank | Where-Object Разница -ne 0)
if ($diff.Count -ne $expected.bank_diffs.Count) { throw 'Different discrepancy count' }
foreach ($day in $expected.bank_diffs) {
    $actual=@($diff | Where-Object 'Дата продажи' -eq $day.date)
    if ($actual.Count -ne 1 -or $actual[0].Разница -ne [decimal]$day.diff) { throw "Daily difference mismatch: $($day.date)" }
}
foreach ($month in $expected.cash) {
    $actual=@($r.Cash | Where-Object Месяц -eq $month.month)
    if ($actual.Count -ne 1 -or $actual[0].Разница -ne [decimal]$month.diff) { throw "Cash mismatch: $($month.month)" }
}
if ($r.Special.Count -ne 22 -or (Sum-Field $r.Special Amount) -ne [decimal]4490.02) { throw 'Expense returns mismatch' }
if ($r.Shifted.Count -ne 1 -or $r.Shifted[0].Amount -ne 6875) { throw 'Settlement timing mismatch' }
if ($r.Controls.Count -ne 8 -or @($r.Controls | Where-Object { $null -eq $_.Difference -or $_.Difference -ne 0 }).Count) { throw 'Printed totals do not reconcile' }
if ($beforeK -ne (Get-FileHash -LiteralPath $kpath).Hash -or $beforeO -ne (Get-FileHash -LiteralPath $opath).Hash) { throw 'Source modified' }
$report=Write-AutoReport $r (Join-Path $PSScriptRoot 'Результат проверки реальных файлов')
$r | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'test-output/actual-real.json') -Encoding UTF8
Write-Output "PASS: both actual workbooks read in Excel, source hashes unchanged; all 18 bank differences and 7 cash months agree with independent Python calculation. Report: $report"
$r.Controls | Format-Table Source,Ref,Label,Value | Out-String | Write-Output
