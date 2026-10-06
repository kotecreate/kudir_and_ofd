$ErrorActionPreference='Stop'
. "$PSScriptRoot/compare.ps1" -NoGui
function Check($ok,$name) { if (-not $ok) { throw "FAILED: $name" } }
function Row($n,$cells) { [pscustomobject]@{Number=$n;Cells=$cells} }
$k1=[pscustomobject]@{Path='synthetic-kudir.xlsx';Sheet='Квартал-1';Rows=@(
    (Row 4 @{B='Дата и номер первичного документа';D='Доходы'}),
    (Row 6 @{A=1;B='31.03.2025';C='Зачисление по QR коду TEST от 31.03.2025';D=50}),
    (Row 7 @{A=2;B='31.03.2025';C='Оплата наличными с 31.03.2025 по 31.03.2025';D=100}),
    (Row 8 @{A='Итого за I квартал';D=150})
)}
$k2=[pscustomobject]@{Path='synthetic-kudir.xlsx';Sheet='Квартал-2';Rows=@(
    (Row 4 @{B='Дата и номер первичного документа';D='Доходы'}),
    (Row 6 @{A=1;B='01.04.2025';C='Зачисление по терминалу TEST за 2025-03-31';D=920}),
    (Row 7 @{A='Итого за II квартал';D=920})
)}
$o=[pscustomobject]@{Path='synthetic-ofd.xlsx';Sheet='Лист1';Date1904=$false;Rows=@(
    (Row 3 @{F='с 31.03.2025';G='по 01.04.2025'}),
    (Row 7 @{K='Дата создания документа';Q='Тип операции';R='Сумма документа, руб.';S='Сумма наличными, руб.';T='Сумма безналичными, руб.'}),
    (Row 8 @{K='31.03.2025';N='Кассовый чек';Q='Приход';R=50;S=0;T=50;D='TEST-FN';J=1}),
    (Row 9 @{K='31.03.2025';N='Кассовый чек';Q='Приход';R=1000;S=100;T=900;D='TEST-FN';J=2}),
    (Row 10 @{K='01.04.2025';N='Кассовый чек';Q='Возврат расхода';R=25;S=0;T=25;D='TEST-FN';J=3}),
    (Row 11 @{Q='Итого всего:';R=1075})
)}
$r=Get-AutoComparison @($k1,$k2) $o
Check ($r.Issues.Count -eq 0) 'Clean synthetic input'
Check ($r.KAll -eq 1070 -and $r.OSales -eq 1050 -and $r.Difference -eq -20) 'Comparable totals'
Check ($r.Special.Count -eq 1 -and $r.OAll -eq 1075) 'Expense return excluded'
$day=@($r.Bank | Where-Object 'Дата продажи' -eq '2025-03-31')[0]
$ex=Get-BankExplanation $r $day
Check ($ex.Target -eq 'Квартал-2!6' -and $ex.Conditional -eq 900) 'Conditional explanation'
Check ($ex.Cards[0].DocDate -eq '2025-04-01') 'Next day settlement'
Check (@($r.Controls | Where-Object Difference -ne 0).Count -eq 0) 'Printed totals'
$o.Rows[2].Cells.Q='Возврат прихода'
$changed=Get-AutoComparison @($k1,$k2) $o
Check ($changed.OSales -eq 950) 'Return sign'
$o.Rows[2].Cells.Q='Unknown'
$changed=Get-AutoComparison @($k1,$k2) $o
Check ($changed.Issues.Count -gt 0) 'Unknown operation surfaced'
$o.Rows[2].Cells.Q='Приход'
$o.Rows[3].Cells.J=1
$changed=Get-AutoComparison @($k1,$k2) $o
Check (@($changed.Issues | Where-Object Причина -like '*повторный*').Count -gt 0) 'Duplicate document detected'
$o.Rows[3].Cells.J=2
$report=Write-AutoReport $r (Join-Path $PSScriptRoot 'test-output/auto')
$html=[IO.File]::ReadAllText($report)
Check ($html.Contains('day-2025-03-31') -and $html.Contains('white-space:nowrap') -and $html.Contains('дата документа в B6')) 'Report dates, pointers and formatting'
Write-Output 'PASS: synthetic automatic reconciliation and explanations'
