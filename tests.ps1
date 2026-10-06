$ErrorActionPreference='Stop'
. "$PSScriptRoot/compare.ps1" -NoGui
$script:checks=0
function Assert($ok,$message) { if (-not $ok) { throw "FAILED: $message" }; $script:checks++ }
function Must-Fail([scriptblock]$action,$message) { $failed=$false; try { & $action | Out-Null } catch { $failed=$true }; Assert $failed $message }
Assert ((Money '1 234,56 руб.') -eq [decimal]1234.56) 'Russian currency'
Assert ((Money '(123,45)') -eq [decimal]-123.45) 'Negative parentheses'
Assert ((Money 0) -eq 0) 'Zero is present'
Assert ($null -eq (Money '')) 'Blank is missing'
Must-Fail { Money '1,234.56' } 'Ambiguous currency rejected'
Must-Fail { Money ([int]-2146826281) } 'Excel error rejected'
Assert ((Read-Date 'Документ № 3 от 03.09.2026').ToString('yyyy-MM-dd') -eq '2026-09-03') 'Embedded date'
Assert ((Read-Date '2026-09-03 11:50:00').Day -eq 3) 'ISO datetime'
Assert ((Read-Date ([datetime]'2026-09-03').ToOADate()).Day -eq 3) 'Excel serial'
Must-Fail { Read-Date '01.09.2026 - 30.09.2026' } 'Two dates rejected'
Must-Fail { Read-Date '31.02.2026' } 'Invalid date rejected'
Must-Fail { Read-Date '01.09.26' } 'Short year rejected'
Assert ((Column-Name 28) -eq 'AB') 'Column naming'
$data=[pscustomobject]@{Path='test.xlsx';Sheet='Sheet';Date1904=$false;Rows=@(
    [pscustomobject]@{Number=1;Cells=@{A='Дата';B='Сумма'}},
    [pscustomobject]@{Number=2;Cells=@{A='03.09.2026';B='100,10'}},
    [pscustomobject]@{Number=3;Cells=@{A='03.09.2026';B='50,20';C='10'}},
    [pscustomobject]@{Number=4;Cells=@{A='04.09.2026';B='0'}},
    [pscustomobject]@{Number=5;Cells=@{A='Итого';B='150,30'}},
    [pscustomobject]@{Number=6;Cells=@{A='плохая дата';B='20'}},
    [pscustomobject]@{Number=7;Cells=@{A='05.09.2026';C='-20'}},
    [pscustomobject]@{Number=8;Cells=@{A='06.09.2026';B='25';D='Итого за день'}}
)}
$k=Extract-Records $data A B C 2 8
Assert ($k.Records.Count -eq 4) 'Valid records count'
Assert (@($k.Audit | Where-Object Level -eq 'Проверить').Count -eq 1) 'Bad date audit'
Assert (@($k.Audit | Where-Object Reason -eq 'Строка итогов или остатка, не включена').Count -eq 2) 'Summary rows excluded including dated total'
Assert ($k.Records[-1].Amount -eq -20) 'Refund only row'
$o=[pscustomobject]@{Records=@(
    [pscustomobject]@{Date='2026-09-03';Amount=[decimal]150.30;Row=2},
    [pscustomobject]@{Date='2026-09-04';Amount=[decimal]0;Row=3},
    [pscustomobject]@{Date='2026-09-06';Amount=[decimal]30;Row=4}
);Audit=@();Data=$data;DateCol='A';AmountCol='B';RefundCol='';First=2;Last=4}
$days=Compare-Records $k.Records $o.Records
Assert ($days.Count -eq 4) 'Union of dates'
Assert ($days[0].Difference -eq 10) 'Same day aggregation and refund'
Assert ($days[0].KRows -eq '2, 3') 'Source trace'
Assert ($days[1].Status -eq 'Суммы за день совпали') 'Zero day present'
Assert ($days[2].Status -eq 'Нет записей ОФД за дату') 'Only KUDIR date'
Assert ($days[3].Status -eq 'Нет записей КУДиР за дату') 'Only OFD date'
Must-Fail { Extract-Records $data A A '' 2 8 } 'Invalid column mapping'
$data1904=[pscustomobject]@{Path='test';Sheet='test';Date1904=$true;Rows=@([pscustomobject]@{Number=1;Cells=@{A=([datetime]'2026-09-03').ToOADate()-1462;B=1}})}
Assert ((Extract-Records $data1904 A B '' 1 1).Records[0].Date -eq '2026-09-03') '1904 date system'
$folder=Join-Path $PSScriptRoot 'test-output'
$report=Write-Report $k $o $folder
$html=[IO.File]::ReadAllText($report)
Assert ($html.Contains('Сверка неполная')) 'Incomplete status'
Assert ($html.Contains('140,30')) 'KUDIR day total'
Assert ((Test-Path (Join-Path $folder 'records.csv'))) 'Report artifacts'
Write-CsvSafe (Join-Path $folder 'escaping.csv') @([pscustomobject]@{Value='=1+1'},[pscustomobject]@{Value='text;"quoted"'}) @('Value')
$csv=[IO.File]::ReadAllText((Join-Path $folder 'escaping.csv'))
Assert ($csv.Contains("'=1+1")) 'CSV formula injection escaped'
Assert ($csv.Contains('""quoted""')) 'CSV quoting'
Assert ((Html '<script>') -eq '&lt;script&gt;') 'HTML escaped'
Write-Output "PASS: $script:checks checks"
