$ErrorActionPreference='Stop'
. "$PSScriptRoot/compare.ps1" -NoGui
$data=Import-Clixml -LiteralPath (Join-Path $PSScriptRoot 'test-output/real-inputs.xml')
$original=Get-AutoComparison $data.K $data.O
$row=@($data.K | Where-Object Sheet -eq 'Квартал-4')[0].Rows | Where-Object Number -eq 6
$saved=$row.Cells['D']; $row.Cells['D']=[double]$saved+100
$changed=Get-AutoComparison $data.K $data.O
if ($changed.Difference -ne $original.Difference-100) { throw 'Changed settlement not detected' }
if (@($changed.Issues | Where-Object Причина -like 'Печатный итог*').Count -ne 2) { throw 'Printed quarter/year total mismatch not detected' }
$row.Cells['D']=$saved
$receipt=@($data.O.Rows | Where-Object { $_.Cells['Q'] -eq 'Приход' })[0]
$saved=$receipt.Cells['Q']; $receipt.Cells['Q']='Возврат прихода'
$changed=Get-AutoComparison $data.K $data.O
if ($changed.OSales -ne $original.OSales-2*(Money $receipt.Cells['R'])) { throw 'Sale return sign wrong' }
$receipt.Cells['Q']='Unknown'
$changed=Get-AutoComparison $data.K $data.O
if (-not @($changed.Issues | Where-Object Причина -like 'Неизвестный тип*').Count) { throw 'Unknown operation silently accepted' }
$receipt.Cells['Q']=$saved
$second=@($data.O.Rows | Where-Object { $_.Cells['Q'] -eq 'Приход' })[1]
$saved=$second.Cells['J']; $second.Cells['J']=$receipt.Cells['J']
$changed=Get-AutoComparison $data.K $data.O
if (-not @($changed.Issues | Where-Object Причина -like 'Пустой или повторный*').Count) { throw 'Duplicate fiscal document not detected' }
$second.Cells['J']=$saved
Write-Output 'PASS: actual-data mutations: amount change, printed controls, sales return, unknown operation, duplicate fiscal document'
