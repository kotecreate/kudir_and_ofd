$ErrorActionPreference='Stop'
. "$PSScriptRoot/compare.ps1" -NoGui
$r=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-output/actual-real.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$day=@($r.Bank | Where-Object 'Дата продажи' -eq '2026-03-31')[0]
$e=Get-BankExplanation $r $day
if ($e.K.Count -ne 3 -or $e.Cards.Count -ne 1 -or $e.QR.Count -ne 2) { throw 'Wrong March 31 decomposition' }
if ($e.Target -ne 'Квартал-2!8' -or $e.Cards[0].DocDate -ne '2026-04-01') { throw 'Wrong settlement pointer' }
if ($e.Conditional -ne 5790 -or $e.Cards[0].Amount -ne 6070) { throw 'Wrong conditional arithmetic' }
if (@($e.O | Where-Object Bank -eq 0).Count) { throw 'Cash-only receipt leaked into bank detail' }
$specialDay=@($r.Bank | Where-Object 'Дата продажи' -eq '2026-09-26')[0]
$se=Get-BankExplanation $r $specialDay
if ($se.Special.Count -ne 1 -or $null -ne $se.Conditional -or $se.First -notlike '*типы*') { throw 'Expense-return warning not prioritized' }
# With several settlements, do not nominate an arbitrary one or invent a correction.
$extra=$e.Cards[0].PSObject.Copy(); $extra.Ref='Квартал-2!999'
$oldK=$r.K; $r.K=@($r.K)+@($extra)
$many=Get-BankExplanation $r $day
if ($many.Target -ne '' -or $null -ne $many.Conditional) { throw 'False certainty with multiple settlements' }
$r.K=$oldK
$report=Write-AutoReport $r (Join-Path $PSScriptRoot 'Результат проверки реальных файлов')
$html=[IO.File]::ReadAllText($report)
foreach ($text in @('white-space:nowrap','min-width:110px','overflow-x:auto','day-2026-03-31','дата документа в B8','5 790,00','не готовая сумма для исправления','Столбец R')) {
    if (-not $html.Replace([char]0xa0,' ').Contains($text)) { throw "Missing explanation or format: $text" }
}
Write-Output 'PASS: date mapping, exact D8 pointer, conditional 5790, cash-only exclusion, expense-return priority, ambiguous settlements, report safeguards'
